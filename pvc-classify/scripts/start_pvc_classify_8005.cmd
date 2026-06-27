@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\pvc-classify\deploy\docker"
docker compose -f docker-compose.yml up -d --build
echo PVC 分类服务 8005 已启动
