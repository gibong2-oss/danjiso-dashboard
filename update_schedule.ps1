# 작업 스케줄러 일정 변경: 매주 월요일 오전 10시
# 기존 작업 찾기 (run_dashboard.bat 를 실행하는 작업)

$taskFound = $false

Get-ScheduledTask | ForEach-Object {
    $task = $_
    $actions = $task.Actions
    foreach ($action in $actions) {
        if ($action.Execute -like "*run_dashboard.bat*" -or 
            ($action.Execute -like "*cmd*" -and $action.Arguments -like "*run_dashboard.bat*")) {
            $taskFound = $true
            $taskName = $task.TaskName
            $taskPath = $task.TaskPath
            
            Write-Host "작업 발견: $taskPath$taskName"
            Write-Host "일정 변경 중: 매주 월요일 오전 10:00..."
            
            # 새 트리거 설정: 매주 월요일 10:00
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "10:00AM"
            
            Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $trigger
            
            Write-Host "완료! 스케줄 변경 완료: 매주 월요일 오전 10:00"
        }
    }
}

if (-not $taskFound) {
    Write-Host "run_dashboard.bat 를 실행하는 작업을 찾지 못했습니다."
    Write-Host ""
    Write-Host "등록된 작업 목록:"
    Get-ScheduledTask | Where-Object { $_.TaskPath -notlike "\Microsoft\*" } | 
        Select-Object TaskName, TaskPath | Format-Table -AutoSize
}

Read-Host "`n완료. Enter 키로 닫기"
