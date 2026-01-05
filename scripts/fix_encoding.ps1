# 读取文件为字节
$filePath = "c:\Users\M1133\excel-copilot-addin\src\taskpane\components\App.tsx"
$bytes = [System.IO.File]::ReadAllBytes($filePath)

# 转为字符串 (UTF-8)
$content = [System.Text.Encoding]::UTF8.GetString($bytes)

# 查找并替换乱码字符串
# 乱码 "正在执行..." 和 "思考中..."
$pattern1 = '����ִ��...'  # 乱码的 "正在执行..."
$pattern2 = '˼����...'    # 乱码的 "思考中..."

if ($content.Contains($pattern1)) {
    $content = $content.Replace($pattern1, '正在执行...')
    Write-Host "Fixed pattern1"
}

if ($content.Contains($pattern2)) {
    $content = $content.Replace($pattern2, '思考中...')
    Write-Host "Fixed pattern2"
}

# 修复更多乱码注释
$commentFixes = @{
    '/* ===== �������� ===== */' = '/* ===== 输入区域 ===== */'
    '/* ===== ��Ϣ�б� ===== */' = '/* ===== 消息列表 ===== */'
    '/* ������ť */' = '/* 操作按钮 */'
    'Excel ��������' = 'Excel 智能助手'
}

foreach ($old in $commentFixes.Keys) {
    if ($content.Contains($old)) {
        $content = $content.Replace($old, $commentFixes[$old])
        Write-Host "Fixed comment: $old"
    }
}

# 保存为UTF-8 with BOM
[System.IO.File]::WriteAllText($filePath, $content, [System.Text.Encoding]::UTF8)
Write-Host "Done"
