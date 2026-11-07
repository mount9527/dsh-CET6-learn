#!/usr/bin/env node
/**
 * 题库快照（打包已构建好的 data/）
 * ---------------------------------------------------------------------------
 * 为什么需要：题库是**本机构建产物**，不在 Git 里，也不在 npm 包里。
 * 换机器、重装系统、或者想回滚到「上一版题库」时，光有资料目录还得重跑一次构建
 * （还要读一堆 docx/pdf）。所以给它做一份带清单的快照，几秒钟就能拷走。
 *
 * 快照里有什么：
 *   index.json / papers/*.json / vocab-*.json / media.json / build-report.json
 * 快照里没有什么：
 *   录音文件本身（几百 MB，且仍在资料目录里）、.cache/（PDF 抽取缓存）。
 *   media.json 记着录音在资料目录里的相对路径，换机器时把录音一起带走即可。
 *
 * 用法：
 *   npm run pack:data                      # 写到 ~/.dsh/dsh-cet6/packs/cet6-<时间>.tgz
 *   npm run pack:data -- --out D:/备份     # 换个位置
 *   npm run pack:data -- --name 2023-2026  # 换个名字前缀
 */
import { execFileSync } from 'node:child_process'
import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { resolveDataDir } from '../lib/data-dir.js'

const argOut = (() => {
  const i = process.argv.indexOf('--out')
  return i > -1 ? process.argv[i + 1] : ''
})()
const argName = (() => {
  const i = process.argv.indexOf('--name')
  return i > -1 ? process.argv[i + 1] : ''
})()

const { dir: DATA, source } = resolveDataDir()
if (!fs.existsSync(path.join(DATA, 'index.json'))) {
  console.error(`✗ 题库还没生成：${DATA} 里没有 index.json`)
  console.error('  先用你自备的资料构建一次：npm run build:data -- --src "D:/你的/CET-6"')
  process.exit(1)
}

const index = JSON.parse(fs.readFileSync(path.join(DATA, 'index.json'), 'utf8'))
const stamp = (() => {
  const d = new Date()
  const p = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`
})()
const slug = (argName || 'cet6').replace(/[^\w.-]+/g, '-')
const outDir = path.resolve(argOut || path.join(path.dirname(DATA), 'packs'))
fs.mkdirSync(outDir, { recursive: true })
const outFile = path.join(outDir, `${slug}-${stamp}.tgz`)

/** 不把 .cache / 录音带进快照：录音在资料目录里，按 media.json 找 */
const skip = [path.join(DATA, '.cache')]
const files = []
;(function walk(dir) {
  for (const d of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, d.name)
    if (skip.some((s) => p === s || p.startsWith(s + path.sep))) continue
    if (d.isDirectory()) walk(p)
    else files.push(p)
  }
})(DATA)

const payloadBytes = files.reduce((n, f) => n + fs.statSync(f).size, 0)

// 用系统自带的 tar 打包：Windows 10+ 与 Linux / macOS 都有，不引第三方依赖。
// stdio 用 ignore 而不是 pipe —— 受限环境（沙箱）里管道式 spawn 会被拒绝。
try {
  execFileSync('tar', ['-czf', outFile, '-C', DATA, '.'], { stdio: 'ignore' })
} catch (e) {
  console.error(`✗ 打包失败（${e.code || e.message}）：找不到可用的 tar，或目标目录不可写`)
  console.error(`  可以自己压：tar -czf "${outFile}" -C "${DATA}" .`)
  process.exit(1)
}

const sha256 = crypto.createHash('sha256').update(fs.readFileSync(outFile)).digest('hex')
const bytes = fs.statSync(outFile).size
const manifest = {
  name: path.basename(outFile, '.tgz'),
  createdAt: new Date().toISOString(),
  from: { dir: DATA, source },
  builtAt: index.generatedAt || null,
  builtFrom: index.source || null,
  papers: (index.papers || []).length,
  vocab: index.vocab || null,
  audio: index.audio || null,
  files: files.length,
  payloadBytes,
  archiveBytes: bytes,
  sha256,
  note: '录音文件不在此快照内；media.json 记着它们在资料目录里的相对路径。',
}
fs.writeFileSync(`${outFile}.json`, JSON.stringify(manifest, null, 1))

const mb = (n) => `${(n / 1024 / 1024).toFixed(1)} MB`
console.log(`题库快照：${outFile}`)
console.log(`  卷宗 ${manifest.papers} 套 · 文件 ${manifest.files} 个 · 原始 ${mb(payloadBytes)} → 压缩 ${mb(bytes)}`)
console.log(`  清单：${outFile}.json（含 sha256）`)
console.log(`  重建：tar -xzf "${path.basename(outFile)}" -C ~/.dsh/dsh-cet6/data`)
