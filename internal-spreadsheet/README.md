# 内部在线表格系统

一个安全、私有的在线协作文档系统，专为替代腾讯/金山等外部在线文档服务而设计，确保企业数据安全。

## 特性

- 🔒 **数据安全**: 所有数据存储在本地服务器，不经过第三方
- 👥 **实时协作**: 支持 30+ 人同时在线编辑，实时同步
- 📊 **电子表格**: 类似 Excel 的在线表格功能
- 🔐 **用户认证**: 完善的登录和权限管理
- 📤 **数据导出**: 支持导出为 CSV 格式
- 🎨 **简洁界面**: 现代化、易用的用户界面

## 技术栈

- **后端**: Python Flask + Flask-SocketIO
- **前端**: HTML5 + CSS3 + JavaScript
- **数据库**: SQLite (可升级为 PostgreSQL)
- **实时通信**: WebSocket (Socket.IO)

## 快速开始

### 1. 安装依赖

```bash
cd internal-spreadsheet
pip install -r requirements.txt
```

### 2. 启动服务

```bash
python backend/app.py
```

### 3. 访问系统

打开浏览器访问: http://localhost:5000

### 4. 默认账号

- 用户名：`admin`
- 密码：`admin123`

**⚠️ 首次登录后请立即修改密码！**

## 目录结构

```
internal-spreadsheet/
├── backend/
│   └── app.py              # Flask 后端应用
├── frontend/
│   └── templates/
│       ├── login.html      # 登录页面
│       ├── index.html      # 首页（表格列表）
│       └── spreadsheet.html # 表格编辑页面
├── data/                    # 数据存储目录
├── requirements.txt         # Python 依赖
└── README.md               # 本说明文件
```

## 使用说明

### 创建表格

1. 登录后点击"+ 新建表格"按钮
2. 输入表格名称
3. 点击"创建"按钮

### 编辑表格

- 点击任意单元格进行编辑
- 使用方向键或 Tab 键在单元格间移动
- Enter 键确认并移动到下一行
- 数据自动保存并实时同步给其他用户

### 导出数据

- 点击工具栏的"导出数据"按钮
- 表格将下载为 CSV 文件

## 部署建议

### 生产环境配置

1. **修改密钥**: 在 `app.py` 中修改 `SECRET_KEY`
2. **数据库升级**: 将 SQLite 改为 PostgreSQL
3. **HTTPS**: 使用 Nginx 反向代理并配置 SSL
4. **多进程**: 使用 Gunicorn 或 uWSGI 运行

### Docker 部署 (可选)

```dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000
CMD ["python", "backend/app.py"]
```

## 安全性

- ✅ 数据完全存储在内部服务器
- ✅ 用户密码加密存储 (bcrypt)
- ✅ Session 管理和 CSRF 保护
- ✅ 实时协作时的数据一致性保证

## 性能指标

- 支持 30+ 并发用户
- 单元格更新延迟 < 100ms
- 单表格支持 50 行 × 26 列 (可扩展)

## 扩展功能建议

- [ ] 公式计算支持
- [ ] 更多导出格式 (Excel, PDF)
- [ ] 表格模板
- [ ] 历史记录和版本控制
- [ ] 更细粒度的权限管理
- [ ] 评论和批注功能

## 故障排除

### 无法启动服务

检查端口是否被占用:
```bash
lsof -i :5000
```

### 数据库错误

删除并重新创建数据库:
```bash
rm backend/instance/spreadsheet.db
python backend/app.py
```

## 技术支持

如有问题，请联系内部 IT 支持团队。

---

**注意**: 本系统为内部使用，请勿部署到公网环境。
