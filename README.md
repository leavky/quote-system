# 报价系统

基于 React、HeroUI 和 Tauri 的本地报价处理工具，包含：

- 项目价格匹配
- 报价匹配
- 价格维护

## 本地开发

```bash
cd 项目结算价格匹配/desktop
npm ci
npm run desktop:dev
```

## Windows 构建

GitHub Actions 的 `Build Windows` 工作流支持手动运行，也会在推送 `v*` 标签时自动运行。构建产物包含免安装程序和 NSIS 安装包。

仓库不包含任何业务 Excel 数据。首次运行时请在界面中选择本机的台账、结算价格表和报价库。
