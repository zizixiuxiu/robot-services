@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\order-split\deploy\docker"
docker compose -f docker-compose.yml up -d
echo 料单拆分服务 8002 已启动
