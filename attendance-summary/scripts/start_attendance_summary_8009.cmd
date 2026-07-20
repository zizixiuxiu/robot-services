@echo off
chcp 65001 >nul
cd /d "D:\Services\robot-services\attendance-summary\deploy\docker"
docker compose -f docker-compose.yml up -d
echo 考勤汇总服务 8009 已启动
