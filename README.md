# 考证闪卡 · 卡片内容

App（[JapanITFlashCard](https://github.com/HHHeart08/JapanITFlashCard)）启动时直接从这个仓库 main 分支的原始文件（raw.githubusercontent.com）读 `manifest.json`，版本比本机新就下载 `cards.json` 和 `catalog.json`。

更新内容：在 App 仓库里改好数据后运行

    python3 tools/publish_content.py ../JapanITFlashCard-data --notes "这次更新了什么"

再把这个仓库推上来即可。版本号会自动加 1。

## Excel 备份和更新记录

`backup/考证闪卡内容备份.xlsx` 是全部卡片内容的 Excel 备份（科目汇总、全部卡片、按科目、更新记录）。
每次更新内容后运行

    python3 tools/export_excel.py --notes "这次更新了什么"

脚本会和上次导出的内容（`backup/snapshot.json`）逐张比较，把新增 / 修改 / 删除追加到
`backup/changelog.json`，再重新生成 Excel。`publish_content.py` 发布完会自动运行它。
