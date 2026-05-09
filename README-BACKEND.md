# WiseAudit 后端部署说明

## 文件清单
- `server.js` - Node代理层（Express）
- `app.py` - Python FastAPI后端
- `audit_core.py` - 审计核心计算引擎（从原项目复制）
- `requirements.txt` - Python依赖
- `package.json` - Node依赖

## 部署步骤
1. 将原项目的 `audit_core.py` 复制到此目录
2. 确保 `requirements.txt` 包含所有依赖
3. 部署到Render（见render.yaml）

## 环境变量
- `DEEPSEEK_API_KEY` - DeepSeek API密钥（必填）
- `PORT` - Node服务端口（默认3001）
- `PYTHON_API_URL` - Python服务地址（默认http://127.0.0.1:8000）
