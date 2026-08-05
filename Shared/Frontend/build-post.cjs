// 构建后处理脚本：移除 HTML 中的 type="module"
const fs = require('fs');
const path = require('path');

// 处理 index.html
const htmlPath = path.join(__dirname, 'wwwroot/index.html');
if (fs.existsSync(htmlPath)) {
  let html = fs.readFileSync(htmlPath, 'utf8');
  
  // 移除 modulepreload 链接（可能导致 CORS 问题）
  html = html.replace(/<link rel="modulepreload"[^>]*>/g, '');
  
  // 保留 type="module"，但移除 crossorigin（WebView2 中不需要）
  html = html.replace(/crossorigin/g, '');
  html = html.replace(/<link rel="stylesheet"\s+crossorigin/g, '<link rel="stylesheet"');
  
  // 确保脚本标签有 type="module"
  html = html.replace(/<script\s+src=/g, '<script type="module" src=');
  
  // 将脚本从 head 移到 body 末尾
  const scriptMatch = html.match(/<script[^>]*src="[^"]*"[^>]*><\/script>/);
  if (scriptMatch) {
    const scriptTag = scriptMatch[0];
    // 从 head 中移除
    html = html.replace(scriptTag, '');
    // 添加到 body 末尾
    html = html.replace('</body>', `  ${scriptTag}\n</body>`);
  }
  
  // 修复可能的引号问题
  html = html.replace(/rel="stylesheet href=/g, 'rel="stylesheet" href=');
  
  fs.writeFileSync(htmlPath, html, 'utf8');
  console.log('✓ 已处理 index.html');
}

// 处理 knowledge-base.html：保留 Vite 注入的 CSS 链接，不要用源码 HTML 覆盖
// （旧逻辑会硬编码不存在的 knowledgeBaseApi.css，导致 Element Plus 样式丢失、图标巨大）
const knowledgeBaseDestPath = path.join(__dirname, 'wwwroot/knowledge-base.html');

if (fs.existsSync(knowledgeBaseDestPath)) {
  let kbHtml = fs.readFileSync(knowledgeBaseDestPath, 'utf8');

  // 移除 modulepreload（可能导致 CORS 问题）
  kbHtml = kbHtml.replace(/<link rel="modulepreload"[^>]*>/g, '');

  // 保留 type="module"，但移除 crossorigin（WebView2 中不需要）
  kbHtml = kbHtml.replace(/crossorigin/g, '');
  kbHtml = kbHtml.replace(/<link rel="stylesheet"\s+crossorigin/g, '<link rel="stylesheet"');

  // 开发入口路径兜底（正常应由 Vite 改写为 ./assets/knowledge-base.js）
  kbHtml = kbHtml.replace(/src="\/src\/knowledge-base-main\.js"/g, 'src="./assets/knowledge-base.js"');
  kbHtml = kbHtml.replace(/src="\.\/assets\/knowledge-base-main\.js"/g, 'src="./assets/knowledge-base.js"');

  // 确保脚本标签有 type="module"
  kbHtml = kbHtml.replace(/<script\s+src=/g, '<script type="module" src=');

  // 将脚本从 head 移到 body 末尾
  const kbScriptMatch = kbHtml.match(/<script[^>]*src="[^"]*"[^>]*><\/script>/);
  if (kbScriptMatch) {
    const scriptTag = kbScriptMatch[0];
    kbHtml = kbHtml.replace(scriptTag, '');
    kbHtml = kbHtml.replace('</body>', `  ${scriptTag}\n</body>`);
  }

  // 修复可能的引号问题
  kbHtml = kbHtml.replace(/rel="stylesheet href=/g, 'rel="stylesheet" href=');

  // 兜底：若 Vite 未注入任何 stylesheet，按实际产物补齐（含 Element Plus 所在的 reset.css）
  if (!kbHtml.includes('<link rel="stylesheet"')) {
    const assetsDir = path.join(__dirname, 'wwwroot/assets');
    const links = [];
    if (fs.existsSync(path.join(assetsDir, 'reset.css'))) {
      links.push('  <link rel="stylesheet" href="./assets/reset.css">');
    }
    if (fs.existsSync(path.join(assetsDir, 'knowledge-base.css'))) {
      links.push('  <link rel="stylesheet" href="./assets/knowledge-base.css">');
    }
    if (links.length) {
      kbHtml = kbHtml.replace('</head>', `${links.join('\n')}\n</head>`);
    }
  }

  fs.writeFileSync(knowledgeBaseDestPath, kbHtml, 'utf8');
  console.log('✓ 已处理 knowledge-base.html（保留 Vite CSS 链接）');
} else {
  console.warn('⚠ 未找到 wwwroot/knowledge-base.html，跳过处理');
}

// 复制并处理 settings.html
const settingsSourcePath = path.join(__dirname, 'settings.html');
const settingsDestPath = path.join(__dirname, 'wwwroot/settings.html');

if (fs.existsSync(settingsSourcePath)) {
  fs.copyFileSync(settingsSourcePath, settingsDestPath);
  console.log('✓ 已复制 settings.html');

  let settingsHtml = fs.readFileSync(settingsDestPath, 'utf8');

  settingsHtml = settingsHtml.replace(/src="\/src\/settings-main\.js"/g, 'src="./assets/settings.js"');
  settingsHtml = settingsHtml.replace(/src="\.\/assets\/settings-main\.js"/g, 'src="./assets/settings.js"');
  settingsHtml = settingsHtml.replace(/crossorigin/g, '');
  settingsHtml = settingsHtml.replace(/<link rel="modulepreload"\s+crossorigin/g, '<link rel="modulepreload"');
  settingsHtml = settingsHtml.replace(/<link rel="stylesheet"\s+crossorigin/g, '<link rel="stylesheet"');
  settingsHtml = settingsHtml.replace(/<script\s+src=/g, '<script type="module" src=');

  if (!settingsHtml.includes('<link rel="stylesheet"')) {
    settingsHtml = settingsHtml.replace('</head>', '  <link rel="stylesheet" href="./assets/settings.css">\n</head>');
  }

  fs.writeFileSync(settingsDestPath, settingsHtml, 'utf8');
  console.log('✓ 已处理 settings.html');
} else {
  console.warn('⚠ 未找到 settings.html，跳过处理');
}

// 复制并处理 login.html
const loginSourcePath = path.join(__dirname, 'login.html');
const loginDestPath = path.join(__dirname, 'wwwroot/login.html');

if (fs.existsSync(loginSourcePath)) {
  fs.copyFileSync(loginSourcePath, loginDestPath);
  console.log('✓ 已复制 login.html');

  let loginHtml = fs.readFileSync(loginDestPath, 'utf8');

  loginHtml = loginHtml.replace(/src="\/src\/login-main\.js"/g, 'src="./assets/login.js"');
  loginHtml = loginHtml.replace(/src="\.\/assets\/login-main\.js"/g, 'src="./assets/login.js"');
  loginHtml = loginHtml.replace(/crossorigin/g, '');
  loginHtml = loginHtml.replace(/<link rel="modulepreload"\s+crossorigin/g, '<link rel="modulepreload"');
  loginHtml = loginHtml.replace(/<link rel="stylesheet"\s+crossorigin/g, '<link rel="stylesheet"');
  loginHtml = loginHtml.replace(/<script\s+src=/g, '<script type="module" src=');

  // 登录页只用 login.css，移除误注入的 Element Plus style.css
  loginHtml = loginHtml.replace(/<link rel="stylesheet"[^>]*href="\.\/assets\/style\.css"[^>]*>\s*/g, '');

  // 确保 login.css 链接存在
  if (!loginHtml.includes('./assets/login.css')) {
    loginHtml = loginHtml.replace('</head>', '  <link rel="stylesheet" href="./assets/login.css">\n</head>');
  }

  fs.writeFileSync(loginDestPath, loginHtml, 'utf8');
  console.log('✓ 已处理 login.html');
} else {
  console.warn('⚠ 未找到 login.html，跳过处理');
}

