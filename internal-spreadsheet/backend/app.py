from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_socketio import SocketIO, emit, join_room, leave_room
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import json
import os

# 获取项目根目录
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
DB_PATH = os.path.join(DATA_DIR, 'spreadsheet.db')

# 确保 data 目录存在
os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__, template_folder='../frontend/templates', static_folder='../frontend/static')
app.config['SECRET_KEY'] = 'internal-spreadsheet-secret-key-2024'
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
socketio = SocketIO(app, cors_allowed_origins="*")
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# 数据库模型
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Spreadsheet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    owner_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    owner = db.relationship('User', backref=db.backref('spreadsheets', lazy=True))

class CellData(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    spreadsheet_id = db.Column(db.Integer, db.ForeignKey('spreadsheet.id'), nullable=False)
    row = db.Column(db.Integer, nullable=False)
    col = db.Column(db.Integer, nullable=False)
    value = db.Column(db.Text, default='')
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    spreadsheet = db.relationship('Spreadsheet', backref=db.backref('cells', lazy=True))
    
    __table_args__ = (db.UniqueConstraint('spreadsheet_id', 'row', 'col', name='unique_cell'),)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# 路由
@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            return jsonify({'success': True})
        return jsonify({'success': False, 'message': '用户名或密码错误'})
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return jsonify({'success': True})

@app.route('/api/spreadsheets', methods=['GET'])
@login_required
def get_spreadsheets():
    spreadsheets = Spreadsheet.query.all()
    return jsonify([{
        'id': s.id,
        'name': s.name,
        'owner': s.owner.username,
        'created_at': s.created_at.isoformat(),
        'updated_at': s.updated_at.isoformat()
    } for s in spreadsheets])

@app.route('/api/spreadsheet', methods=['POST'])
@login_required
def create_spreadsheet():
    data = request.json
    name = data.get('name')
    if not name:
        return jsonify({'success': False, 'message': '表格名称不能为空'}), 400
    
    spreadsheet = Spreadsheet(name=name, owner_id=current_user.id)
    db.session.add(spreadsheet)
    db.session.commit()
    
    # 初始化一些单元格
    for row in range(1, 21):
        for col in range(1, 11):
            cell = CellData(spreadsheet_id=spreadsheet.id, row=row, col=col, value='')
            db.session.add(cell)
    db.session.commit()
    
    return jsonify({'success': True, 'id': spreadsheet.id})

@app.route('/spreadsheet/<int:id>')
@login_required
def view_spreadsheet(id):
    spreadsheet = Spreadsheet.query.get_or_404(id)
    return render_template('spreadsheet.html', spreadsheet_id=id, spreadsheet_name=spreadsheet.name)

@app.route('/api/spreadsheet/<int:id>/data', methods=['GET'])
@login_required
def get_spreadsheet_data(id):
    cells = CellData.query.filter_by(spreadsheet_id=id).all()
    data = {}
    for cell in cells:
        key = f"{cell.row},{cell.col}"
        data[key] = cell.value
    return jsonify(data)

@app.route('/api/user', methods=['GET'])
@login_required
def get_current_user():
    return jsonify({
        'id': current_user.id,
        'username': current_user.username
    })

# SocketIO 事件
@socketio.on('connect')
def handle_connect():
    print(f'用户连接: {request.sid}')

@socketio.on('join')
def handle_join(data):
    spreadsheet_id = data.get('spreadsheet_id')
    room = f'spreadsheet_{spreadsheet_id}'
    join_room(room)
    print(f'用户 {request.sid} 加入房间 {room}')

@socketio.on('cell_update')
def handle_cell_update(data):
    spreadsheet_id = data.get('spreadsheet_id')
    row = data.get('row')
    col = data.get('col')
    value = data.get('value')
    user_id = data.get('user_id')
    
    # 更新数据库
    cell = CellData.query.filter_by(
        spreadsheet_id=spreadsheet_id, 
        row=row, 
        col=col
    ).first()
    
    if cell:
        cell.value = value
        cell.updated_at = datetime.utcnow()
        db.session.commit()
    else:
        cell = CellData(
            spreadsheet_id=spreadsheet_id,
            row=row,
            col=col,
            value=value
        )
        db.session.add(cell)
        db.session.commit()
    
    # 广播给其他用户
    room = f'spreadsheet_{spreadsheet_id}'
    emit('cell_updated', {
        'row': row,
        'col': col,
        'value': value,
        'user_id': user_id
    }, room=room, skip_sid=request.sid)

@socketio.on('disconnect')
def handle_disconnect():
    print(f'用户断开连接: {request.sid}')

def init_db():
    with app.app_context():
        db.create_all()
        
        # 创建默认管理员用户
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(
                username='admin',
                password_hash=generate_password_hash('admin123')
            )
            db.session.add(admin)
            db.session.commit()
            print('默认管理员用户已创建: admin / admin123')

if __name__ == '__main__':
    init_db()
    socketio.run(app, host='0.0.0.0', port=5000, debug=True, allow_unsafe_werkzeug=True)
