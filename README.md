# Feishu Project Deliverables Plugin

飞书项目交付物插件原型，支持 DOCX 交付物上传、版本保存、基线管理、附件阅览/下载、版本 Diff 与变更导航。

## Local Setup

1. Install dependencies:

```bash
npm install
```

2. Create local plugin config:

```bash
cp plugin.config.example.json plugin.config.json
```

3. Fill `plugin.config.json` with your own `pluginId` and `pluginSecret`.

4. Start Feishu Project plugin debugging:

```bash
lpm --cwd "$(pwd)" start --auto
```

## Scripts

```bash
npm test
```

## Security

`plugin.config.json` contains the plugin secret and is intentionally ignored by Git. Do not commit real plugin credentials to a public repository.

`package-lock.json` is also ignored because the generated lockfile may contain internal registry URLs.
