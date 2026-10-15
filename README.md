# Z40

赵诗琳的 DSH 技能仓库 —— 一个地方存放所有自定义 skill，装到电脑上后 DSH 每次对话自动识别调用。

## 里面有什么

| Skill | 干什么用的 |
|---|---|
| [`desktop-control`](skills/desktop-control/) | 让 AI 能操作 Windows 桌面：找窗口、点按钮、填表单、读屏幕内容、截图。靠无障碍元素定位，不靠鼠标坐标（坐标点击在这台电脑上无效） |
| [`ppt-master`](skills/ppt-master/) | AI 生成可编辑的 PPTX 演示文稿。也支持重建页面视觉、做品牌/样式/版式模板、填充原生 PPTX 模板、美化成品 PPT、加旁白和动画 |

## 一键安装到本机

```powershell
git clone https://github.com/zhaoshilin40/Z40.git D:\ai\Z40
cd D:\ai\Z40
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

装完**重启 DSH Desktop**，之后每次对话 AI 都能看到这些 skill，需要时自动调用。

## 它是怎么工作的

DSH 启动时扫描 `C:\Users\<用户名>\.dsh\skills\` 目录，发现里面的 skill 并把摘要注入对话。
`install.ps1` 不复制文件，而是给每个 skill 建一个**目录联接（junction）**：

```
C:\Users\<用户名>\.dsh\skills\ppt-master   ──junction──▶   D:\ai\Z40\skills\ppt-master
```

好处是**改一处两边同步**：

- 在这里 `git pull` 拉到新版本 → 电脑上立刻生效，不用重装
- AI 在电脑上改了 skill 文件 → 在仓库里 `git add` 就能提交回 GitHub

卸载只需要删掉 `.dsh\skills\` 里对应的联接，仓库本身不受影响。

## 加一个新 skill

1. 在 `skills/` 下建目录，放一个 `SKILL.md`，开头必须有 YAML frontmatter：

   ```markdown
   ---
   name: my-skill
   description: 一句话说明什么时候该用这个 skill
   ---

   # 正文：写给 AI 的操作说明
   ```

2. 跑一遍 `install.ps1`（它会跳过已存在的），或手动建联接：

   ```powershell
   New-Item -ItemType Junction `
     -Path "$env:USERPROFILE\.dsh\skills\my-skill" `
     -Target "D:\ai\Z40\skills\my-skill"
   ```

3. 重启 DSH Desktop。

## 注意

- **密钥不要提交。** `.gitignore` 已排除 `.env`，但提交前仍建议 `git status` 看一眼。
  ppt-master 的配图 API key 放在 `C:\Users\<用户名>\.ppt-master\.env`，不在这个仓库里。
- 仓库是 **私有** 的，只有自己能看。
- `ppt-master` 体积较大（约 37 MB，主要是图标和音效模板），克隆会稍慢。
