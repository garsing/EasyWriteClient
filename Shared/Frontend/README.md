# EasyWrite Shared Frontend

Desktop 与 Word Plugin 共用的 Vue 前端。

```bash
cd Shared/Frontend
npm install
npm run build
```

产物输出到本目录 `wwwroot/`。

- Desktop：编译时拷贝到输出目录；运行时也可直接解析本路径
- Plugin：编译时同步到 `Plugin/wwwroot`
