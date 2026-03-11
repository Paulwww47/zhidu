#!/bin/bash
cd "$(dirname "$0")"
source venv/Scripts/activate
echo "========================================="
echo "  制度编写系统启动中..."
echo "========================================="
echo ""
echo "  编辑页面: http://127.0.0.1:5000"
echo "  管理后台: http://127.0.0.1:5000/admin/login"
echo "  默认管理员: admin / admin123"
echo ""
echo "  按 Ctrl+C 停止服务"
echo "========================================="
echo ""
python app.py
