@echo off
chcp 65001 >nul
echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║     康恩贝行业信息简报 - 高级数据看板 v4.2          ║
echo ╠══════════════════════════════════════════════════════╣
echo ║  更新: 支持12家信息来源（11家竞品+国家药监局）      ║
echo ║  特性: 词云图 / 数据图表 / 自动刷新 / 智能分析     ║
echo ╚══════════════════════════════════════════════════════╝
echo.

:: 安装依赖
echo [1/3] 正在安装依赖组件...
pip install streamlit==1.32.2 streamlit-autorefresh==1.0.1 plotly==5.18.0 wordcloud==1.9.3 matplotlib==3.8.0 -q

:: 获取邮件数据
echo.
echo [2/3] 正在获取最新邮件...
python get_emails.py

:: 启动看板
echo.
echo [3/3] 正在启动高级看板...
echo.
echo ══════════════════════════════════════════════════════
echo   浏览器将自动打开，如未打开请访问:
echo   http://localhost:8501
echo.
echo   按 Ctrl+C 可停止服务
echo ══════════════════════════════════════════════════════
echo.
streamlit run app_dashboard.py --server.port 8501
