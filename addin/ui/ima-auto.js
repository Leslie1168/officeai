/* Shared by the WebView UI and dependency-injected regression tests. */
const ImaAuto = (() => {
  const normalize = s => String(s || '').toLowerCase().replace(/\.(md|txt|docx|html?)$/i, '').replace(/[\s·._《》“”"-]/g, '');
  const excluded = text => /(?:不要|不用|无需|不必|别|不)(?:再)?(?:去)?(?:参考|查阅|查询|搜索|检索|使用|访问|读取)?\s*ima/i.test(text);
  const requested = text => /\bima\b/i.test(text) && !excluded(text);
  function parse(raw) {
    if (typeof raw !== 'string') return raw;
    const clean = raw.replace(/^\s*```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');
    return JSON.parse(clean);
  }
  async function retrieve(text, api, progress, active) {
    const guard = () => { if (!active()) throw new Error('检索已停止'); };
    const call = async (method, ...args) => { guard(); const value = parse(await api[method](...args)); guard(); return value; };
    progress('正在连接 IMA，读取知识库列表…');
    const listed = await call('ImaListKb', '');
    if (!listed.ok) throw new Error('IMA 连接失败：' + (listed.message || '未知错误'));
    const libraries = (listed.items || []).filter(k => k.id);
    if (!libraries.length) throw new Error('IMA 没有可访问的知识库。');
    progress('正在识别要查的知识库和资料名…');
    const quoted = [...text.matchAll(/《([^》]+)》/g)].map(m => m[1].trim());
    const namedLibraries = libraries.filter(k => k.name && text.includes(k.name));
    // A fully specified target does not need model interpretation.
    const plan = quoted.length && quoted.length <= 3 && namedLibraries.length ? {
      kbIds:namedLibraries.map(k => k.id), titles:quoted,
      queries:[...new Set(quoted.map(t => t.replace(/\.(md|txt|docx|html?)$/i, '')))]
    } : await call('ImaPlan', text, JSON.stringify(libraries.map(({id,name,baseType}) => ({id,name,baseType}))));
    if (plan.ok === false) throw new Error('IMA 检索计划失败：' + plan.message);
    if (plan.clarification) throw new Error(plan.clarification);
    if (!Array.isArray(plan.kbIds) || !Array.isArray(plan.queries)) throw new Error('未能识别检索目标，请说明知识库和资料名称。');
    const selected = libraries.filter(k => plan.kbIds.includes(k.id)).slice(0, 3);
    const queries = [...new Set(plan.queries.filter(q => typeof q === 'string' && q.trim()).map(q => q.trim().slice(0, 80)))].slice(0, 3);
    const titles = Array.isArray(plan.titles) ? plan.titles.filter(t => typeof t === 'string' && t.trim()).slice(0, 3) : [];
    if (!selected.length || !queries.length) throw new Error('请指定要参考的 IMA 知识库或资料名，例如：参考 ima 小说库里的《元衍纪·设定》。');
    const found = new Map(); const failures = [];
    for (const kb of selected) {
      for (const query of queries) {
        progress(`正在搜索「${kb.name}」：${query}…`);
        const result = await call('ImaSearch', kb.id, query);
        if (!result.ok) { failures.push(kb.name + '：' + (result.message || '搜索失败')); continue; }
        for (const item of result.items || []) if (item.mediaId) found.set(item.mediaId, {...item, kbName:kb.name});
      }
    }
    let candidates = [...found.values()];
    const exact = t => candidates.filter(i => normalize(i.title) === normalize(t));
    if (titles.length) {
      const missing = titles.filter(t => !exact(t).length);
      if (missing.length) throw new Error('未找到指定资料：' + missing.join('、') + '。请核对标题，或只提供资料名关键词。' + (failures.length ? '\n' + failures.join('\n') : ''));
      candidates = titles.flatMap(exact);
    }
    const score = item => queries.reduce((n,q) => n + (normalize(item.title).includes(normalize(q)) ? 10 : 0), 0) + (/\.(md|txt)$/i.test(item.title) ? 2 : 0);
    candidates.sort((a,b) => score(b)-score(a));
    candidates = candidates.filter((item, index, all) => all.findIndex(other => normalize(other.title) === normalize(item.title)) === index);
    if (!candidates.length) throw new Error('IMA 未找到相关资料。试试资料标题关键词，例如“元衍”“设定”。' + (failures.length ? '\n' + failures.join('\n') : ''));
    const refs = []; const unreadable = []; let size = 0;
    for (const item of candidates.slice(0, 6)) {
      if (refs.length >= 3) break;
      progress(`正在读取《${item.title}》…`);
      const media = await call('ImaGetMedia', item.mediaId);
      if (!media.ok || !media.content || /^(PK\u0003\u0004|%PDF-)/.test(media.content)) { unreadable.push(item.title); continue; }
      const content = media.content.slice(0, Math.min(20000, 60000-size));
      if (!content) break;
      size += content.length;
      refs.push({mediaId:item.mediaId,title:item.title,kbName:item.kbName,content,auto:true});
    }
    guard();
    const missedTitles = titles.filter(t => !refs.some(r => normalize(r.title) === normalize(t)));
    if (!refs.length || missedTitles.length) throw new Error('资料已找到，但未能读取所需正文：' + (missedTitles.length ? missedTitles : unreadable).join('、') + '。可读取 TXT、Markdown、DOCX；本次不会据此猜测回答。');
    return {refs, summary:'IMA 已检索并引用：\n' + refs.map(r => `• ${r.kbName} / ${r.title}（已读取 ${r.content.length} 字，单篇最多 2 万字）`).join('\n') +
      (failures.length ? '\n部分检索失败：' + failures.join('；') : '') + (unreadable.length ? '\n未读取：' + unreadable.join('、') : '')};
  }
  return {requested,excluded,retrieve};
})();
if (typeof module !== 'undefined') module.exports = ImaAuto;
