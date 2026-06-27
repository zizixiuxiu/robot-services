@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\csv-board\deploy\docker"
docker compose -f docker-compose.yml up -d --build
echo CSV 板件转换服务 8004 已启动
