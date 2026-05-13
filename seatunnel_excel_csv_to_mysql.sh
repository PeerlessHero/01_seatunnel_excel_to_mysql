#!/bin/bash

# SeaTunnel Excel/CSV to MySQL 批量处理脚本
# 版本: v0.1
# 作者: ETL Team
# 描述: 递归遍历目录，批量处理 Excel 和 CSV 文件，将数据同步到 MySQL

# 配置参数
SEATUNNEL_HOME="/data/ksf/seatunnel/apache-seatunnel-2.3.13"
MYSQL_HOST="192.168.203.128"
MYSQL_PORT="3306"
MYSQL_USER="etl"
MYSQL_PASS="Etl@1234"
MYSQL_DB="mid"
MYSQL_LOG_TABLE="excel_file_sync_log"
CONF_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXCEL_CONF="${CONF_DIR}/seatunnel_excel_to_mysql_v0.1.conf"
CSV_CONF="${CONF_DIR}/seatunnel_csv_to_mysql_v0.1.conf"
LOG_FILE="${CONF_DIR}/seatunnel_sync_$(date +%Y%m%d_%H%M%S).log"
TEMP_CONF_DIR="${CONF_DIR}/temp_confs_$(date +%Y%m%d_%H%M%S)"

# 创建临时配置文件目录
mkdir -p "$TEMP_CONF_DIR"
log_info "创建临时配置文件目录: $TEMP_CONF_DIR"

# 检查参数
if [ $# -ne 1 ]; then
    echo "使用方法: $0 <目录路径>"
    echo "示例: $0 /data/excel_files"
    exit 1
fi

ROOT_DIR="$1"

# 检查目录是否存在
if [ ! -d "$ROOT_DIR" ]; then
    echo "错误: 目录 $ROOT_DIR 不存在" | tee -a "$LOG_FILE"
    exit 1
fi

# 检查 SeaTunnel 安装路径
if [ ! -d "$SEATUNNEL_HOME" ]; then
    echo "错误: SeaTunnel 安装路径 $SEATUNNEL_HOME 不存在" | tee -a "$LOG_FILE"
    exit 1
fi

# 检查配置文件是否存在
if [ ! -f "$EXCEL_CONF" ]; then
    echo "错误: Excel 配置文件 $EXCEL_CONF 不存在" | tee -a "$LOG_FILE"
    exit 1
fi

if [ ! -f "$CSV_CONF" ]; then
    echo "错误: CSV 配置文件 $CSV_CONF 不存在" | tee -a "$LOG_FILE"
    exit 1
fi

# 日志函数
log_info() {
    local msg="[$(date '+%Y-%m-%d %H:%M:%S')] INFO: $1"
    echo "$msg" | tee -a "$LOG_FILE"
}

log_error() {
    local msg="[$(date '+%Y-%m-%d %H:%M:%S')] ERROR: $1"
    echo "$msg" | tee -a "$LOG_FILE"
}

log_success() {
    local msg="[$(date '+%Y-%m-%d %H:%M:%S')] SUCCESS: $1"
    echo "$msg" | tee -a "$LOG_FILE"
}

# 插入日志到 MySQL
insert_log() {
    local sync_time="$1"
    local file_path="$2"
    local file_name="$3"
    local sheet_name="$4"
    local status="$5"
    local total_rows="$6"
    local error_msg="$7"
    local job_name="$8"

    # 确保 total_rows 有默认值
    if [ -z "$total_rows" ] || ! [[ "$total_rows" =~ ^[0-9]+$ ]]; then
        total_rows=0
    fi

    # 转义特殊字符（改进版）
    # 使用 printf 进行更安全的转义
    error_msg=$(printf '%s' "$error_msg" | sed -e 's/\\/\\\\/g' -e "s/'/''/g" -e 's/"/\\"/g' | tr '\n' ' ' | tr '\r' ' ' | tr '\t' ' ')
    file_path=$(printf '%s' "$file_path" | sed -e 's/\\/\\\\/g' -e "s/'/''/g" -e 's/"/\\"/g')
    file_name=$(printf '%s' "$file_name" | sed -e 's/\\/\\\\/g' -e "s/'/''/g" -e 's/"/\\"/g')
    sheet_name=$(printf '%s' "$sheet_name" | sed -e 's/\\/\\\\/g' -e "s/'/''/g" -e 's/"/\\"/g')
    job_name=$(printf '%s' "$job_name" | sed -e 's/\\/\\\\/g' -e "s/'/''/g" -e 's/"/\\"/g')

    # 构建 SQL 语句（使用双引号转义单引号）
    local sql="INSERT INTO $MYSQL_LOG_TABLE (sync_time, file_path, file_name, sheet_name, status, total_rows, error_msg, job_name) VALUES ('$sync_time', '$file_path', '$file_name', '$sheet_name', '$status', $total_rows, '$error_msg', '$job_name');"

    # 打印完整的 MySQL 命令（便于调试）
    local mysql_cmd="mysql -h$MYSQL_HOST -P$MYSQL_PORT -u$MYSQL_USER -p*** --default-character-set=utf8mb3 $MYSQL_DB -e \"$sql\""
    log_info "执行 MySQL 命令: $mysql_cmd"

    # 执行 MySQL 命令，设置字符集为 utf8mb3
    local mysql_output
    mysql_output=$(mysql -h"$MYSQL_HOST" -P"$MYSQL_PORT" -u"$MYSQL_USER" -p"$MYSQL_PASS" --default-character-set=utf8mb3 "$MYSQL_DB" -e "$sql" 2>&1)
    local mysql_exit_code=$?

    if [ $mysql_exit_code -ne 0 ]; then
        log_error "插入日志到 MySQL 失败: $file_path - $sheet_name"
        log_error "MySQL 错误: $mysql_output"
        log_error "SQL: $sql"
        return 1
    fi

    return 0
}

# 获取 Excel 文件的 sheet 名称
get_excel_sheets() {
    local excel_file="$1"
    local sheets=()

    # 使用 Python 获取 sheet 名称
    if command -v python3 &> /dev/null; then
        python3 << EOF 2>/dev/null
import sys
try:
    import openpyxl
    wb = openpyxl.load_workbook('$excel_file', read_only=True)
    for sheet_name in wb.sheetnames:
        print(sheet_name)
    wb.close()
except Exception as e:
    sys.exit(1)
EOF
        if [ $? -eq 0 ]; then
            return 0
        fi
    fi

    # 如果 openpyxl 不可用，尝试使用 xlrd
    if command -v python3 &> /dev/null; then
        python3 << EOF 2>/dev/null
import sys
try:
    import xlrd
    wb = xlrd.open_workbook('$excel_file', on_demand=True)
    for sheet_name in wb.sheet_names():
        print(sheet_name)
    wb.release_resources()
except Exception as e:
    sys.exit(1)
EOF
        if [ $? -eq 0 ]; then
            return 0
        fi
    fi

    # 如果 Python 库都不可用，尝试使用 ssconvert (gnumeric)
    if command -v ssconvert &> /dev/null; then
        ssconvert --list-exporters "$excel_file" 2>/dev/null | grep -oP 'Sheet\d+:\s*\K.*' || echo "Sheet1"
        return 0
    fi

    # 如果都不可用，返回默认的 sheet 名称
    echo "Sheet1"
    return 0
}

# 处理 Excel 文件
process_excel() {
    local excel_file="$1"
    local file_name=$(basename "$excel_file")
    local file_ext="${file_name##*.}"
    local sync_time=$(date '+%Y-%m-%d %H:%M:%S')
    local job_name="excel_to_mysql_$(date +%Y%m%d_%H%M%S)"

    log_info "开始处理 Excel 文件: $excel_file"

    # 获取所有 sheet 名称
    local sheets=$(get_excel_sheets "$excel_file")
    if [ -z "$sheets" ]; then
        log_error "无法获取 Excel 文件的 sheet 信息: $excel_file"
        insert_log "$sync_time" "$excel_file" "$file_name" "" "fail" 0 "无法获取 sheet 信息" "$job_name"
        return 1
    fi

    # 处理每个 sheet
    local sheet_count=0
    local success_count=0
    local fail_count=0

    while IFS= read -r sheet_name; do
        if [ -z "$sheet_name" ]; then
            continue
        fi

        sheet_count=$((sheet_count + 1))
        log_info "  处理 Sheet: $sheet_name"

        # 创建临时配置文件（使用文件名称和数字编号命名）
        local temp_conf="${TEMP_CONF_DIR}/temp_${file_name}_${sheet_count}.conf"
        
        # 使用 sed 将所有占位符替换为实际值
        sed "s|\${excel_path}|$excel_file|g; s|\${sheet_name}|$sheet_name|g" "$EXCEL_CONF" > "$temp_conf"

        # 执行 SeaTunnel 任务
        local start_time=$(date +%s)
        local seatunnel_cmd="$SEATUNNEL_HOME/bin/seatunnel.sh --config $temp_conf -e local"
        log_info "执行 SeaTunnel 命令: $seatunnel_cmd"
        local output=$($seatunnel_cmd 2>&1)
        local exit_code=$?
        local end_time=$(date +%s)
        local duration=$((end_time - start_time))

        # 打印 SeaTunnel 输出（便于调试）
        log_info "SeaTunnel 输出: $output"
        log_info "SeaTunnel 退出码: $exit_code"

        # 获取处理行数
        local total_rows=0
        if [ $exit_code -eq 0 ]; then
            # 从输出中提取行数信息
            total_rows=$(echo "$output" | grep -oP 'Total rows: \K\d+' || echo "0")
            if [ "$total_rows" = "0" ]; then
                # 尝试其他方式获取行数
                total_rows=$(echo "$output" | grep -oP 'rows: \K\d+' | head -1 || echo "0")
            fi
        fi

        # 检查输出中是否包含错误信息
        local has_error=false
        if echo "$output" | grep -qiE 'error|exception|failed|fail|traceback' > /dev/null 2>&1; then
            has_error=true
            log_info "检测到输出中包含错误信息"
        fi

        # 临时配置文件已保存到: $temp_conf

        # 记录结果（同时检查退出码和错误信息）
        if [ $exit_code -eq 0 ] && [ "$has_error" = false ]; then
            log_success "  Sheet $sheet_name 处理成功，耗时 ${duration}s，行数: $total_rows"
            insert_log "$sync_time" "$excel_file" "$file_name" "$sheet_name" "success" "$total_rows" "" "$job_name"
            success_count=$((success_count + 1))
        else
            local error_msg=$(echo "$output" | tail -50 | tr '\n' ' ' | sed 's/"/\\"/g')
            log_error "  Sheet $sheet_name 处理失败: $error_msg"
            insert_log "$sync_time" "$excel_file" "$file_name" "$sheet_name" "fail" "$total_rows" "$error_msg" "$job_name"
            fail_count=$((fail_count + 1))
        fi
    done <<< "$sheets"

    log_info "Excel 文件 $excel_file 处理完成: 总 Sheet 数 $sheet_count, 成功 $success_count, 失败 $fail_count"
    return 0
}

# 处理 CSV/DAT 文件
process_csv() {
    local csv_file="$1"
    local file_name=$(basename "$csv_file")
    local file_ext="${file_name##*.}"
    local sync_time=$(date '+%Y-%m-%d %H:%M:%S')
    local job_name="csv_to_mysql_$(date +%Y%m%d_%H%M%S)"

    log_info "开始处理 CSV/DAT 文件: $csv_file"

    # 创建临时配置文件（使用文件名命名）
    local temp_conf="${TEMP_CONF_DIR}/temp_${file_name}.conf"
    
    # 使用 sed 将所有占位符替换为实际值
    sed "s|\${csv_path}|$csv_file|g" "$CSV_CONF" > "$temp_conf"

    # 执行 SeaTunnel 任务
    local start_time=$(date +%s)
    local seatunnel_cmd="$SEATUNNEL_HOME/bin/seatunnel.sh --config $temp_conf -e local"
    log_info "执行 SeaTunnel 命令: $seatunnel_cmd"
    local output=$($seatunnel_cmd 2>&1)
    local exit_code=$?
    local end_time=$(date +%s)
    local duration=$((end_time - start_time))

    # 打印 SeaTunnel 输出（便于调试）
    log_info "SeaTunnel 输出: $output"
    log_info "SeaTunnel 退出码: $exit_code"

    # 获取处理行数
    local total_rows=0
    if [ $exit_code -eq 0 ]; then
        total_rows=$(echo "$output" | grep -oP 'Total rows: \K\d+' || echo "0")
        if [ "$total_rows" = "0" ]; then
            total_rows=$(echo "$output" | grep -oP 'rows: \K\d+' | head -1 || echo "0")
        fi
    fi

    # 检查输出中是否包含错误信息
    local has_error=false
    if echo "$output" | grep -qiE 'error|exception|failed|fail|traceback' > /dev/null 2>&1; then
        has_error=true
        log_info "检测到输出中包含错误信息"
    fi

    # 临时配置文件已保存到: $temp_conf

    # 记录结果（同时检查退出码和错误信息）
    if [ $exit_code -eq 0 ] && [ "$has_error" = false ]; then
        log_success "CSV/DAT 文件 $csv_file 处理成功，耗时 ${duration}s，行数: $total_rows"
        insert_log "$sync_time" "$csv_file" "$file_name" "default" "success" "$total_rows" "" "$job_name"
    else
        local error_msg=$(echo "$output" | tail -50 | tr '\n' ' ' | sed 's/"/\\"/g')
        log_error "CSV/DAT 文件 $csv_file 处理失败: $error_msg"
        insert_log "$sync_time" "$csv_file" "$file_name" "default" "fail" "$total_rows" "$error_msg" "$job_name"
    fi

    return 0
}

# 主处理函数
process_files() {
    local target_dir="$1"
    local total_files=0
    local success_files=0
    local fail_files=0

    log_info "开始处理目录: $target_dir"

    # 查找所有 Excel 和 CSV/DAT 文件
    while IFS= read -r -d '' file; do
        total_files=$((total_files + 1))
        local file_name=$(basename "$file")
        local file_ext="${file_name##*.}"
        local file_ext_lower=$(echo "$file_ext" | tr '[:upper:]' '[:lower:]')

        log_info "发现文件 [$total_files]: $file_name"

        # 根据文件扩展名处理
        case "$file_ext_lower" in
            xlsx|xls)
                if process_excel "$file"; then
                    success_files=$((success_files + 1))
                else
                    fail_files=$((fail_files + 1))
                fi
                ;;
            csv|dat)
                if process_csv "$file"; then
                    success_files=$((success_files + 1))
                else
                    fail_files=$((fail_files + 1))
                fi
                ;;
            *)
                log_info "跳过不支持的文件类型: $file_name"
                ;;
        esac

        # 添加延迟，避免系统负载过高
        sleep 0.5

    done < <(find "$target_dir" -type f \( -iname "*.xlsx" -o -iname "*.xls" -o -iname "*.csv" -o -iname "*.dat" \) -print0)

    log_info "目录处理完成: 总文件数 $total_files, 成功 $success_files, 失败 $fail_files"
    log_info "详细日志请查看: $LOG_FILE"
}

# 主程序
main() {
    log_info "=========================================="
    log_info "SeaTunnel Excel/CSV to MySQL 批量处理脚本启动"
    log_info "版本: v0.1"
    log_info "目标目录: $ROOT_DIR"
    log_info "SeaTunnel 路径: $SEATUNNEL_HOME"
    log_info "MySQL: $MYSQL_HOST:$MYSQL_PORT/$MYSQL_DB"
    log_info "=========================================="

    # 检查 MySQL 连接
    mysql -h"$MYSQL_HOST" -P"$MYSQL_PORT" -u"$MYSQL_USER" -p"$MYSQL_PASS" -e "USE $MYSQL_DB;" 2>/dev/null
    if [ $? -ne 0 ]; then
        log_error "MySQL 连接失败，请检查连接配置"
        exit 1
    fi

    # 检查日志表是否存在
    mysql -h"$MYSQL_HOST" -P"$MYSQL_PORT" -u"$MYSQL_USER" -p"$MYSQL_PASS" "$MYSQL_DB" -e "DESCRIBE $MYSQL_LOG_TABLE;" &>/dev/null
    if [ $? -ne 0 ]; then
        log_error "日志表 $MYSQL_LOG_TABLE 不存在，请先创建该表"
        exit 1
    fi

    # 开始处理文件
    process_files "$ROOT_DIR"

    log_info "=========================================="
    log_info "所有任务完成"
    log_info "=========================================="
}

# 执行主程序
main
