@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\hardware-summary\deploy\docker"
docker compose -f docker-compose.yml up -d
echo 五金汇总服务 8001 已启动
