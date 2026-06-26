@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\workshop-order\deploy\docker"
docker compose -f docker-compose.yml up -d
echo 下车间单服务 8006 已启动
