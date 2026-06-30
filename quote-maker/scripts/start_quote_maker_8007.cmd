@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\quote-maker\deploy\docker"
docker compose -f docker-compose.yml up -d --build
echo 报价单生成服务 8007 已启动
