# 从 GitHub 下载 ZIP 时，下面两个 $Format 会被换成提交号和提交时间（见 .gitattributes 的 export-subst），
# 调试日志里据此知道用户跑的是哪个版本。用 git clone 的目录没有替换，改为现查 git。
COMMIT = "$Format:%h$"
COMMIT_DATE = "$Format:%ci$"
