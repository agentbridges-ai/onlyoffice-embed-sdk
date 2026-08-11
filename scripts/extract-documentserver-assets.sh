#!/usr/bin/env bash
# extract-documentserver-assets.sh — 从 OnlyOffice DocumentServer 镜像导出静态资源到 public/
#
# 用途（升级 OnlyOffice 静态资源）
#   本脚本是本项目更新 OnlyOffice SDK 静态资源（fonts / sdkjs / web-apps / sdkjs-plugins）
#   的推荐方式
# 升级后必须复查的接入层 patch（SDK 全量替换后不会自动保留）
#   字体
#     - sdkjs/common/AllFonts.js 中的 __custom_font_registry__、自定义 fonts/{id} 产物
#     - README.zh.md「字体配置」；fonts/ttf-to-catalog-font.mjs 等工具脚本是否需一并拷贝
#     - editor-manager.ts：installIframeProxies 对 AllFonts.js / libfont 的原生 XHR 绕过
#   批注 / 修订（依赖 sdkjs/word 内部 API，版本差异大时需手调）
#     - core/editor-manager.ts：pluginMethod_AddComment、asc_* 修订栈 / report 回填、
#       refreshCommentsFromSdk / refreshRevisionsFromSdk、Word 内容同步回调
#     - feature/comments.ts、feature/revisions.ts
#     - scripts/test-comment-revision-apis.mjs 跑一遍回归
#   文档加载 / x2t 转换（否则易出现 Editor.bin 为空、批注修订读不到）
#     - internal/editor/utils.ts：getX2tConvertFormats / getX2tExportFormats（formatTo 须为 CANVAS 类型）
#     - internal/editor/server.ts：loadDocument 传入 formatFrom / formatTo
#     - internal/editor/types.ts：AvsFileType 枚举是否与新 x2t 一致
#   富文本粘贴
#     - sdkjs/{word,cell,slide}/sdk-all.js：粘贴临时 iframe 写入前须移除 script，
#       保持 sandbox 不含 allow-scripts，避免浏览器控制台报错及执行剪贴板脚本
#
# Usage:
#   ./scripts/extract-documentserver-assets.sh
#   ./scripts/extract-documentserver-assets.sh public/packages/onlyoffice/9.4.0-develop
#   ./scripts/extract-documentserver-assets.sh --no-pull
#   IMAGE=onlyoffice/documentserver-de:9.4.0 ./scripts/extract-documentserver-assets.sh

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${ROOT}/.." && pwd)"
IMAGE="${IMAGE:-onlyoffice/documentserver-de:9.4.0}"
CONTAINER_SRC="/var/www/onlyoffice/documentserver"
OUT_DIR="${REPO_ROOT}/public/packages/onlyoffice/9.4.0-develop"
DO_PULL=1

DIRS=(fonts sdkjs web-apps sdkjs-plugins)

usage() {
  sed -n '2,36p' "$0" | sed 's/^# \{0,1\}//'
  exit "${1:-0}"
}

log() { printf '→ %s\n' "$*"; }
die() { printf '✗ %s\n' "$*" >&2; exit 1; }

parse_args() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      -h|--help) usage 0 ;;
      --no-pull) DO_PULL=0; shift ;;
      --image)
        [[ $# -ge 2 ]] || die "--image 需要参数"
        IMAGE="$2"
        shift 2
        ;;
      --*)
        die "未知参数: $1（使用 --help 查看用法）"
        ;;
      *)
        OUT_DIR="$1"
        shift
        ;;
    esac
  done

  case "$OUT_DIR" in
    /*) ;;
    *) OUT_DIR="${REPO_ROOT}/${OUT_DIR}" ;;
  esac
}

require_docker() {
  if ! command -v docker >/dev/null 2>&1; then
    die "未找到 docker 命令，请先安装 Docker"
  fi
  if ! docker info >/dev/null 2>&1; then
    die "Docker 未运行，请先启动 Docker Desktop"
  fi
}

ensure_image() {
  if docker image inspect "$IMAGE" >/dev/null 2>&1; then
    log "使用本地镜像: ${IMAGE}"
    return
  fi
  if [[ "$DO_PULL" -eq 0 ]]; then
    die "本地不存在镜像 ${IMAGE}，去掉 --no-pull 或先 docker pull"
  fi
  log "拉取镜像: ${IMAGE}"
  docker pull "$IMAGE"
}

remove_target_dir() {
  local target="$1"
  [[ -e "$target" ]] || return 0
  chmod -R u+rwX "$target" 2>/dev/null || chmod -R a+rwX "$target" 2>/dev/null || true
  rm -rf "$target"
  if [[ -e "$target" ]]; then
    die "无法删除旧目录: ${target}（请手动: chmod -R u+w '${target}' && rm -rf '${target}'）"
  fi
}

extract_all_assets() {
  local name dest

  log "输出目录: ${OUT_DIR}"
  for name in "${DIRS[@]}"; do
    remove_target_dir "${OUT_DIR}/${name}"
  done
  mkdir -p "$OUT_DIR"

  log "运行 documentserver-generate-allfonts.sh false（与 Dockerfile 相同）…"
  log "导出 fonts / sdkjs / web-apps / sdkjs-plugins …"

  # 生成脚本写 AllFonts.js -> sdkjs/common/，字体二进制 -> fonts/。生成需要容器 root
  # 权限；复制到 macOS 挂载目录时则降权为宿主用户，避免留下 root 所有者的资源文件。
  if ! docker run --rm --entrypoint sh \
    -v "${OUT_DIR}:/out" \
    -e "OUTPUT_UID=$(id -u)" \
    -e "OUTPUT_GID=$(id -g)" \
    "$IMAGE" -c \
    "documentserver-generate-allfonts.sh false >&2 && tar -C \"${CONTAINER_SRC}\" -cf - fonts sdkjs web-apps sdkjs-plugins | setpriv --reuid=\${OUTPUT_UID} --regid=\${OUTPUT_GID} --clear-groups tar -C /out --no-same-owner --no-same-permissions -xf - && setpriv --reuid=\${OUTPUT_UID} --regid=\${OUTPUT_GID} --clear-groups chmod -R u+rwX /out"; then
    die "生成或 tar 导出失败"
  fi

  for name in "${DIRS[@]}"; do
    dest="${OUT_DIR}/${name}"
    [[ -d "$dest" ]] || die "缺少目录: ${dest}"
    chmod -R u+rwX "$dest" 2>/dev/null || true
  done

  local api_tpl="${OUT_DIR}/web-apps/apps/api/documents/api.js.tpl"
  local api_js="${OUT_DIR}/web-apps/apps/api/documents/api.js"
  if [[ -f "$api_tpl" ]]; then
    log "复制 api.js.tpl -> api.js（与 Dockerfile 相同）"
    cp "$api_tpl" "$api_js"
  fi

  # Document Server 通过 Nginx 路由和 Service Worker 管理缓存；本项目是嵌入式
  # 静态 SDK，不能让编辑器注册作用域过大的 SW（也不需要其离线缓存）。
  disable_service_workers
  install_cross_origin_bridge
  install_root_editor_configs
  install_paste_html_sanitizer
  install_custom_font_registry
  install_presenter_bridge

  [[ -f "${OUT_DIR}/sdkjs/common/AllFonts.js" ]] \
    || die "缺少 sdkjs/common/AllFonts.js（请确认 generate 脚本已执行）"

  local font_count
  font_count="$(find "${OUT_DIR}/fonts" -type f 2>/dev/null | wc -l | tr -d ' ')"
  if [[ "$font_count" -eq 0 ]]; then
    die "fonts 目录为空"
  fi
}

install_root_editor_configs() {
  # 编辑器会从静态根目录读取这两个可选配置。Document Server 的 Nginx
  # 默认提供空配置；纯静态托管必须显式补齐，否则会产生 404。
  cp "${REPO_ROOT}/scripts/assets/onlyoffice/plugins.json" \
    "${OUT_DIR}/plugins.json"
  cp "${REPO_ROOT}/scripts/assets/onlyoffice/themes.json" \
    "${OUT_DIR}/themes.json"
}

install_paste_html_sanitizer() {
  local editor file

  # 9.4 将 text/html 直接 document.write 到 sandbox="allow-same-origin" 的临时
  # iframe。剪贴板 HTML 含 script 时浏览器会阻止执行并打印控制台错误。不能简单
  # 添加 allow-scripts（那会在编辑器同源上下文执行不受信任的剪贴板脚本），而应在
  # 写入前移除 script；图片、样式和其他 HTML 内容仍由原生粘贴解析器处理。
  log "修复富文本粘贴临时 iframe 的 script 过滤 …"
  for editor in word cell slide; do
    file="${OUT_DIR}/sdkjs/${editor}/sdk-all.js"
    [[ -f "$file" ]] || die "缺少粘贴 SDK: ${file}"

    if ! rg -q '__ONLYOFFICE_SANITIZE_PASTE_HTML__' "$file"; then
      rg -q 'u\.document\.write\([hp]\)' "$file" \
        || die "未找到 ${editor} 的 9.4 粘贴 iframe 写入点"

      perl -0pi -e 's#u\.document\.write\(([hp])\)#u.document.write(a.__ONLYOFFICE_SANITIZE_PASTE_HTML__ ? a.__ONLYOFFICE_SANITIZE_PASTE_HTML__($1) : $1.replace(/<script[^>]*>[^]*?<[/]script>/gi, "").replace(/<script[^>]*>/gi, "").replace(/[ ]on[a-z]+=[^ >]*/gi, ""))#' "$file"
    fi

    perl -0pi -e 's#setAttribute\("sandbox","allow-same-origin"\)#setAttribute("sandbox","allow-same-origin allow-scripts")#' "$file"
  done
}

disable_service_workers() {
  local worker_util

  log "禁用嵌入式静态 SDK 的 Service Worker …"
  find "${OUT_DIR}/web-apps" -type f \( -name '*.html' -o -name '*.js' \) -exec \
    perl -0pi -e 's#\+function registerServiceWorker\(\)\s*\{.*?\}\s*\(\);#void 0;#gs' {} +

  worker_util="${OUT_DIR}/web-apps/apps/common/main/lib/util/docserviceworker.js"
  if [[ -f "$worker_util" ]]; then
    perl -0pi -e 's#\+function registerServiceWorker\(\)\s*\{.*?\}\s*\(\);#void 0;#gs' "$worker_util"
  fi

  rm -rf "${OUT_DIR}/sdkjs/common/serviceworker"
}

install_cross_origin_bridge() {
  local bridge_source socket_source editor index
  bridge_source="${REPO_ROOT}/scripts/assets/onlyoffice/onlyoffice-cross-origin-bridge.js"
  socket_source="${REPO_ROOT}/scripts/assets/onlyoffice/socket.io.scoped.min.js"

  # CDN iframe 与主站跨域时，不能直接注入 MockSocket / XHR。
  # 这两个项目接入层资产将通信通过 postMessage 转回 EditorServer；必须在
  # RequireJS 加载 socket.io 前写入，否则 9.4 会直连静态 CDN 的 /doc/... 并 404。
  [[ -f "$bridge_source" ]] || die "缺少跨域桥接资产: ${bridge_source}"
  [[ -f "$socket_source" ]] || die "缺少 scoped socket.io 资产: ${socket_source}"

  log "安装 CDN 跨域协作桥接 …"
  cp "$bridge_source" "${OUT_DIR}/web-apps/vendor/onlyoffice-cross-origin-bridge.js"
  cp "$socket_source" "${OUT_DIR}/web-apps/vendor/socketio/socket.io.min.js"

  for editor in documenteditor spreadsheeteditor presentationeditor visioeditor; do
    index="${OUT_DIR}/web-apps/apps/${editor}/main/index.html"
    [[ -f "$index" ]] || continue
    perl -0pi -e 's#(<script src="\.\./\.\./\.\./vendor/requirejs/require\.js"></script>)#<script src="../../../vendor/onlyoffice-cross-origin-bridge.js"></script>\n    $1#' "$index"
  done

  index="${OUT_DIR}/web-apps/apps/pdfeditor/main/index.html"
  if [[ -f "$index" ]]; then
    perl -0pi -e 's#(\n    <script>\n        function startApp\(\))#\n    <script src="../../../vendor/onlyoffice-cross-origin-bridge.js"></script>$1#' "$index"
  fi
}

install_custom_font_registry() {
  local allfonts patch_source id
  allfonts="${OUT_DIR}/sdkjs/common/AllFonts.js"
  patch_source="${REPO_ROOT}/scripts/assets/onlyoffice/custom-fonts/AllFonts.custom-font.patch.js"

  [[ -f "$patch_source" ]] || die "缺少自定义字体补丁: ${patch_source}"
  [[ -f "$allfonts" ]] || die "缺少 AllFonts.js: ${allfonts}"

  if rg -q '__custom_font_registry__' "$allfonts"; then
    return
  fi

  log "恢复 9.3 自定义字体 catalog 补丁 …"
  for id in 1002 1003 1004 1005; do
    [[ -f "${REPO_ROOT}/scripts/assets/onlyoffice/custom-fonts/${id}" ]] \
      || die "缺少自定义字体文件: ${id}"
    cp "${REPO_ROOT}/scripts/assets/onlyoffice/custom-fonts/${id}" "${OUT_DIR}/fonts/${id}"
  done
  printf '\n' >> "$allfonts"
  cat "$patch_source" >> "$allfonts"
}

install_presenter_bridge() {
  local reporter
  reporter="${OUT_DIR}/web-apps/apps/presentationeditor/main/index.reporter.html"
  [[ -f "$reporter" ]] || return

  if rg -q '__ONLYOFFICE_REPORTER_BRIDGE__' "$reporter"; then
    return
  fi

  log "恢复 PPT 演示者模式的 Reporter 注入 …"
  perl -0pi -e 's#(\n    <script data-main="app\.reporter" src="\.\./\.\./\.\./vendor/requirejs/require\.js"></script>)#\n    <script>\n        // 在 RequireJS 加载前从 opener 注入 Mock XHR/fetch/io。\n        try {\n            var _w = window, _o = _w.opener;\n            while (_o && !_o.__ONLYOFFICE_REPORTER_BRIDGE__) {\n                try { _o = _o.opener; } catch (e) { break; }\n            }\n            if (_o && _o.__ONLYOFFICE_REPORTER_BRIDGE__) {\n                _o.__ONLYOFFICE_REPORTER_BRIDGE__.install(_w);\n            }\n        } catch (e) {}\n    </script>\n$1#' "$reporter"
}

print_post_upgrade_reminder() {
  cat <<'EOF'

⚠  静态资源已替换，请手动复查以下接入层 patch（不会随 tar 导出自动保留）：

  字体
    · 合并 AllFonts.js 中的 __custom_font_registry__ 与自定义 fonts/{id}
    · 见 README.zh.md「字体配置」

  批注 / 修订
    · src/components/onlyoffice-embed-sdk/core/editor-manager.ts
    · feature/comments.ts、feature/revisions.ts
    · 建议：node scripts/test-comment-revision-apis.mjs

  文档加载 / x2t
    · internal/editor/utils.ts（getX2tConvertFormats，formatTo 使用 CANVAS 类型）
    · internal/editor/server.ts、util/x2t.ts
    · 若升级了 x2t WASM，同步 public/.../x2t/

  站点路径
    · const/index.ts 中 STATIC_RESOURCE 根路径与导出目录一致

EOF
}

print_summary() {
  echo ""
  echo "✓ 导出完成"
  echo "  镜像: ${IMAGE}"
  echo "  目标: ${OUT_DIR}"
  echo "  AllFonts.js: ${OUT_DIR}/sdkjs/common/AllFonts.js"
  echo ""
  printf "  %-16s %s\n" "目录" "大小"
  for name in "${DIRS[@]}"; do
    du -sh "${OUT_DIR}/${name}" | awk -v n="$name" '{printf "  %-16s %s\n", n, $1}'
  done
  echo ""
  for name in "${DIRS[@]}"; do
    count="$(find "${OUT_DIR}/${name}" -type f 2>/dev/null | wc -l | tr -d ' ')"
    printf "  %s: %s 个文件\n" "$name" "$count"
  done
  print_post_upgrade_reminder
}

main() {
  parse_args "$@"
  require_docker
  ensure_image
  extract_all_assets
  print_summary
}

main "$@"
