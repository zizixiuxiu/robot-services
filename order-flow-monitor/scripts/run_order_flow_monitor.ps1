$ErrorActionPreference = "Stop"

$env:ORDER_FLOW_BASE_DIR = "D:\Services\robot-services\order-flow-monitor"
$env:ORDER_FLOW_WORKDIR = "D:\Services\robot-services\order-flow-monitor\src"
$env:ORDER_FLOW_STATE_PATH = "D:\Services\robot-services\order-flow-monitor\data\order_flow_refresh_state.json"
$env:ORDER_FLOW_LOG_PATH = "D:\Services\robot-services\order-flow-monitor\logs\order_flow_refresh.log"
$env:ORDER_FLOW_PYTHON_EXE = "C:\Users\Administrator\.workbuddy\binaries\python\envs\bom-server\Scripts\python.exe"

& $env:ORDER_FLOW_PYTHON_EXE "$env:ORDER_FLOW_WORKDIR\monitor_refresh_order_flow.py" --interval 900
