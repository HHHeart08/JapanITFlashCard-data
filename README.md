# 考证闪卡 · 卡片内容

iOS App「考证闪卡」（[JapanITFlashCard](https://github.com/HHHeart08/JapanITFlashCard)，私有仓库）的公开内容仓库。App 启动时直接读这个仓库 main 分支的原始文件（raw.githubusercontent.com），不需要开 GitHub Pages。

## 文件

| 文件 | 用途 |
|---|---|
| `manifest.json` | 内容版本号和更新说明（现在是版本 2：90 门科目、944 张卡） |
| `cards.json` | 全部名词卡 |
| `catalog.json` | 主办方、科目和免费规则 |
| `privacy-policy.html` / `privacy-policy.md` | 隐私政策（App Store Connect 填的隐私政策网址） |
| `support.html` | 技术支持页（App Store Connect 填的支持网址） |
| `.nojekyll` | 让 GitHub Pages 原样发布 HTML |

`cards.json` 和 `catalog.json` 由 App 仓库的 `tools/publish_content.py` 生成，不要在这里手改，改了也会在下次发布时被覆盖。

## App 怎么更新内容

1. 读 `manifest.json`，`version` 比本机正在用的大才下载。
2. 下载 `cards.json`、`catalog.json`，校验通过后下次打开 App 生效。
3. `manifest.json` 里可以加 `minAppBuild`：构建号低于它的旧版 App 不会下载，用在新数据有旧版 App 不认识的字段时。

只改 `cards.json` / `catalog.json` 而不加版本号，已安装的 App 不会下载。

## 发布新内容

在 App 仓库里改好 `KaozhengFlashcards/Resources/` 下的数据后运行

    python3 tools/publish_content.py ../JapanITFlashCard-data --notes "这次更新了什么"

版本号会自动加 1。然后把这个仓库推到 main。GitHub 原始文件有几分钟缓存，用户下次启动 App 下载，再下次启动生效。

隐私政策改了的话，App 内「设置 → 隐私政策」也要同步改（`Views/InfoViews.swift`）。

联系：hhheartkz@outlook.com
