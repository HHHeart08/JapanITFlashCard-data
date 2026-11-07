# 考证闪卡 · 卡片内容

App（[JapanITFlashCard](https://github.com/HHHeart08/JapanITFlashCard)）启动时从这里的 GitHub Pages 读 `manifest.json`，版本比本机新就下载 `cards.json` 和 `catalog.json`。

更新内容：在 App 仓库里改好数据后运行

    python3 tools/publish_content.py ../JapanITFlashCard-data --notes "这次更新了什么"

再把这个仓库推上来即可。版本号会自动加 1。
