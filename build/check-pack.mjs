#!/usr/bin/env node
/**
 * 数据包体检（构建前先看目录对不对）
 * ---------------------------------------------------------------------------
 * 只读**目录结构与文件名**，不解析 docx / pdf / mp3 内容，所以几秒就能跑完。
 * 它回答的是「我新丢进来的东西，构建能不能认出来」：
 *
 *   - 01–05 五个分类目录找不找得到（允许改掉编号后面的名字）；
 *   - 每套卷的答案、听力原文、录音，按考期 / 套号配得上配不上；
 *   - 词汇 CSV 的列名对不对（构建只认「单词」「释义」这两列）；
 *   - 文件名里能不能解析出 `2024年12月 第2套` 这种考期与套号。
 *
 * 真正的解析质量（选项数、答案数、空号还原）由 build:data 与 check:data 负责，
 * 这里不做重复工作，只把「放错地方 / 起错名字」挡在构建之前。
 *
 * 用法：
 *   node build/check-pack.mjs                 # 默认检查 data-pack/（没有就 ./CET-6）
 *   node build/check-pack.mjs --src "D:/CET-6"
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const argSrc = (() => {
  const i = process.argv.indexOf('--src')
  return i > -1 ? process.argv[i + 1] : ''
})()

/** 默认资料目录：项目里的 data-pack/（有 01 开头的目录才算），否则 ./CET-6 */
function defaultSrc() {
  const pack = path.join(ROOT, 'data-pack')
  try {
    if (fs.readdirSync(pack).some((n) => /^01/.test(n))) return pack
  } catch {
    /* 没有 data-pack/ */
  }
  return path.join(process.cwd(), 'CET-6')
}

const SRC = path.resolve(argSrc || process.env.CET6_SRC || defaultSrc())
const problems = []
const notes = []
const problem = (msg) => problems.push(msg)

const MB = (n) => `${(n / 1024 / 1024).toFixed(1)} MB`

function walk(dir, out = []) {
  let list = []
  try {
    list = fs.readdirSync(dir, { withFileTypes: true })
  } catch {
    return out
  }
  for (const d of list) {
    const p = path.join(dir, d.name)
    if (d.isDirectory()) walk(p, out)
    else out.push(p)
  }
  return out
}

/** 按前缀找分类目录（编号后面的名字可以改；04/05 谁先谁后由内容决定） */
function findDir(prefix) {
  let names = []
  try {
    names = fs.readdirSync(SRC, { withFileTypes: true }).filter((d) => d.isDirectory()).map((d) => d.name)
  } catch {
    return ''
  }
  const hit =
    names.find((n) => n.startsWith(prefix)) ||
    names.find((n) => new RegExp(`(^|[^0-9])${prefix}([^0-9]|$)`).test(n))
  return hit ? path.join(SRC, hit) : ''
}

const pad2 = (n) => String(n).padStart(2, '0')

/** 从文件名解析考期与套号；解析不出返回 null（构建同样会忽略这类文件） */
function keyOf(file) {
  const m = path.basename(file).match(/(\d{4})\s*年\s*(\d{1,2})\s*月.*?第\s*([1-3])\s*套/)
  if (!m) return null
  return { id: `${m[1]}-${pad2(m[2])}-s${m[3]}`, per: `${m[1]}-${pad2(m[2])}` }
}

const firstLine = (file) => {
  try {
    return fs.readFileSync(file, 'utf8').split(/\r?\n/)[0] || ''
  } catch {
    return ''
  }
}

/* ────────────────────────── 开始体检 ────────────────────────── */

if (!fs.existsSync(SRC)) {
  console.error(`✗ 资料目录不存在：${SRC}`)
  console.error('  用 npm run check:pack -- --src "D:/你的/CET-6" 指定，或把资料放进 data-pack/')
  process.exit(1)
}

console.log(`资料目录：${SRC}\n`)

const dir01 = findDir('01')
const dir02 = findDir('02')
const dir03 = findDir('03')
const tailDirs = [findDir('04'), findDir('05')].filter(Boolean)
const dirVocab =
  tailDirs.find((d) => walk(d).some((f) => /\.csv$/i.test(f) && /词/.test(path.basename(f)))) || ''
const dirAudio = tailDirs.find((d) => walk(d).some((f) => /\.(mp3|aac|m4a|wav)$/i.test(f))) || ''
if (!dir01) problem('缺 01 真题试卷 目录（编号开头是识别依据）')
if (!dir02) problem('缺 02 答案 目录')

/* 1. 试卷 */
const papers = new Map() // id → { per, files: [] }
if (dir01) {
  for (const f of walk(dir01).filter((x) => /\.(docx|pdf|txt)$/i.test(x))) {
    const k = keyOf(f)
    if (!k) continue
    if (!papers.has(k.id)) papers.set(k.id, { per: k.per, files: [] })
    papers.get(k.id).files.push(f)
  }
  if (!papers.size) {
    problem('01 真题试卷 里没有一套能解析出「YYYY年M月第N套」的卷子（.docx / .pdf / .txt）')
  }
}

/** 一套卷可能同时有「整卷」「听力真题」「写作真题」：整卷优先，其次 docx → pdf → txt */
function pickWhole(files) {
  const rank = (f) => (/\.docx$/i.test(f) ? 0 : /\.pdf$/i.test(f) ? 1 : 2)
  const whole = files.filter((f) => /整卷/.test(path.basename(f)))
  const pool = whole.length ? whole : files.filter((f) => !/听力|写作|翻译|答案/.test(path.basename(f)))
  return (pool.sort((a, b) => rank(a) - rank(b))[0] || files.sort((a, b) => rank(a) - rank(b))[0])
}

const periods = new Map() // per → Set(id)
for (const [id, v] of papers) {
  if (!periods.has(v.per)) periods.set(v.per, new Set())
  periods.get(v.per).add(id)
}
if (papers.size) {
  const chosen = [...papers.values()].map((v) => pickWhole(v.files))
  const docx = chosen.filter((f) => /\.docx$/i.test(f)).length
  const pdf = chosen.filter((f) => /\.pdf$/i.test(f)).length
  const txt = chosen.length - docx - pdf
  const parts = [`${docx} 套走 docx`]
  if (pdf) parts.push(`${pdf} 套走 pdf`)
  if (txt) parts.push(`${txt} 套是纯文本`)
  console.log(`01 真题试卷   ${periods.size} 个考期 / ${papers.size} 套（${parts.join('、')}）`)
  if (periods.size && papers.size / periods.size !== 3) {
    notes.push(`平均每个考期 ${(papers.size / periods.size).toFixed(1)} 套（常见是 3 套）`)
  }
}

/* 2. 答案 */
const answers = new Map()
if (dir02) {
  for (const f of walk(dir02).filter((x) => x.toLowerCase().endsWith('.txt'))) {
    const k = keyOf(f)
    if (k) answers.set(k.id, f)
    else notes.push(`02 答案 里有解析不出考期的文件：${path.basename(f)}（构建会忽略）`)
  }
  const csv = walk(dir02).find((f) => /\.csv$/i.test(f))
  if (csv) {
    const head = firstLine(csv)
    const has = ['场次', '套数', '题号', '答案'].every((c) => head.includes(c))
    if (!has) problem(`答案总表的表头缺列（需含 场次 / 套数 / 题号 / 答案）：${path.basename(csv)}`)
    else answers.set('__csv__', csv)
  }
  if (!answers.size) problem('02 答案 里既没有单套 TXT，也没有答案总表 CSV')
  if (answers.size) console.log(`02 答案       ${answers.has('__csv__') ? answers.size - 1 : answers.size} 套${answers.has('__csv__') ? '（另有一份 CSV 总表）' : ''}`)
}

/* 3. 听力原文 */
const transcripts = new Map()
if (dir03) {
  for (const f of walk(dir03).filter((x) => x.toLowerCase().endsWith('.txt'))) {
    const k = keyOf(f)
    if (k) transcripts.set(k.id, f)
    else notes.push(`03 听力原文 里的汇总文件没有考期，构建会跳过：${path.basename(f)}`)
  }
  if (transcripts.size) console.log(`03 听力原文   ${transcripts.size} 套`)
  else notes.push('03 听力原文 是空的：听力只能靠听，交卷后不给原文')
} else {
  notes.push('没有 03 听力原文 目录（可选）')
}

/* 4. 听力音频 */
const audios = new Map()
if (dirAudio) {
  let bytes = 0
  for (const f of walk(dirAudio).filter((x) => /\.(mp3|aac|m4a|wav)$/i.test(x))) {
    const k = keyOf(f)
    if (!k) {
      problem(`04 听力音频 的文件名解析不出「YYYY年M月第N套」：${path.basename(f)}`)
      continue
    }
    if (audios.has(k.id)) notes.push(`04 听力音频 里 ${k.id} 有两份，构建只会用其中一份`)
    audios.set(k.id, f)
    bytes += fs.statSync(f).size
  }
  if (audios.size) console.log(`04 听力音频   ${audios.size} 份 / ${MB(bytes)}`)
  else notes.push('04 听力音频 是空的：听力题只能看原文')
} else {
  notes.push(tailDirs.length ? '04/05 里还没有音频文件（可选）' : '没有 04 听力音频 目录（可选）')
}

/* 5. 词汇 */
if (dirVocab) {
  const csvs = walk(dirVocab).filter((f) => /\.csv$/i.test(f))
  const core = csvs.find((f) => /2500/.test(path.basename(f)))
  const full = csvs.find((f) => /6500/.test(path.basename(f))) || csvs.find((f) => /6000/.test(path.basename(f)))
  if (!core) problem('05 词汇 里找不到文件名带 2500 的核心词表（*.csv）')
  if (!full) problem('05 词汇 里找不到文件名带 6500 / 6000 的大纲词表（*.csv）')
  const shape = (f) => {
    const head = firstLine(f)
    const ok = head.includes('单词') && head.includes('释义')
    if (!ok) problem(`词表表头缺「单词」或「释义」列：${path.basename(f)}`)
    const rows = fs.readFileSync(f, 'utf8').split(/\r?\n/).filter((l) => l.trim()).length - 1
    return rows
  }
  const parts = []
  if (core) parts.push(`核心约 ${shape(core)} 词`)
  if (full) parts.push(`大纲约 ${shape(full)} 词`)
  for (const f of csvs) {
    if (f !== core && f !== full) notes.push(`05 词汇 里这张表不会被构建读取：${path.basename(f)}`)
  }
  console.log(`05 词汇       ${csvs.length} 张 CSV（${parts.join(' / ') || '未识别'}）`)
} else {
  problem(
    tailDirs.length
      ? '04/05 里没有 CSV 词表（文件名需含 2500 / 6500 或 6000，表头需有「单词」「释义」）'
      : '缺 05 词汇 目录（核心 2500 与大纲词表都从这里读）'
  )
}

/* 6. 分段表 */
const tracks = [path.join(SRC, 'audio-tracks.json'), dirAudio ? path.join(dirAudio, 'audio-tracks.json') : '']
  .filter(Boolean)
  .find((f) => {
    try {
      return fs.statSync(f).isFile()
    } catch {
      return false
    }
  })
console.log(`分段表       ${tracks ? path.relative(SRC, tracks) : '未提供（听力整段播放，不提供按节跳转）'}`)

/* 7. 覆盖情况 */
if (papers.size) {
  const hasAnswer = [...papers.keys()].filter((id) => answers.has(id) || answers.has('__csv__')).length
  const hasTr = [...papers.keys()].filter((id) => transcripts.has(id)).length
  const hasAu = [...papers.keys()].filter((id) => audios.has(id)).length
  console.log(
    `覆盖          答案 ${hasAnswer}/${papers.size} · 原文 ${hasTr}/${papers.size} · 录音 ${hasAu}/${papers.size}` +
      (hasAu < papers.size ? '（缺的套卷靠同考期借用或只能看原文）' : '')
  )
  const noAnswer = [...papers.keys()].filter((id) => !answers.has(id) && !answers.has('__csv__'))
  if (noAnswer.length) problem(`这些套卷没有答案：${noAnswer.join(', ')}`)
  for (const id of [...papers.keys()].filter((x) => !transcripts.has(x) && audios.has(x))) {
    notes.push(`${id} 有录音但没原文：录音归属只能按文件名推定，界面会提示自己听一遍`)
  }
}

/* 8. 输出 */
console.log('')
if (notes.length) {
  console.log(`提示 ${notes.length} 条：`)
  for (const n of notes) console.log(`  · ${n}`)
}
if (problems.length) {
  console.log(`\n✗ 问题 ${problems.length} 条（先修这些，构建会卡住或结果不对）：`)
  for (const p of problems) console.log(`  · ${p}`)
  console.log('\n改完再跑一次：npm run check:pack')
  process.exit(1)
}

console.log(`✓ 体检通过（问题 0 条${notes.length ? `，提示 ${notes.length} 条` : ''}）`)
console.log('下一步：npm run build:data   （构建完再 npm run check:data）')
