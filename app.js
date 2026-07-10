/**
 * CBIL 단원 설계 도우미 — v3.0 (GitHub Pages 정적 앱)
 * ------------------------------------------------------------------
 * 역할: 기존 앱스크립트 서버(.gs)를 브라우저 안에서 그대로 재현한다.
 *   · google.script.run 호환 심(shim) — 클라이언트(Index.html) 코드는 무수정 이식
 *   · 성취기준/전략은행/통합교과 = repo의 data/*.json (SpreadsheetApp 대체)
 *   · Gemini 호출 = 브라우저 fetch 직접 호출 (UrlFetchApp 대체, 429/503 사다리 유지)
 *   · API 키 = localStorage (UserProperties 대체 — 이 브라우저에만 저장)
 *   · 내보내기 = 다음 빌드에서 .docx 다운로드로 구현 예정(현재 안내 스텁)
 * 두뇌(CBIL 단계 로직)는 v2.5 CBIL.gs 원문 그대로다 — 아래 [CBIL 원문] 구획.
 * ⓒ 영쌤클래스
 */
(function(){
'use strict';

/* ══════════ 0. 설정 ══════════ */
var GEMINI_MODEL          = 'gemini-3.5-flash';
var GEMINI_FALLBACK_MODEL = 'gemini-3.1-flash-lite';
var GEMINI_TEMPERATURE = 0.7;
var GEMINI_MAX_TOKENS  = 8192;
var GEMINI_API_VERSION = 'v1beta';
var KEY_STORE = 'cbil_gemini_key_v1';
var DATA_BASE = 'data/';

/* ══════════ 1. 데이터 계층 (SpreadsheetApp 대체) ══════════ */
var indexCache = null;          // standards-index.json
var bandCache  = {};            // 학년군 → 행 배열(레거시 모양)
var bankCache  = null;          // 전략은행 { rows: [...] }
var tonghapCache = null;        // { units: [...] }

function fetchJson_(name){
  return fetch(DATA_BASE + name, { cache: 'no-cache' }).then(function(res){
    if (!res.ok) throw new Error('데이터 파일을 불러오지 못했습니다(' + name + ', HTTP ' + res.status + '). 저장소의 data 폴더를 확인해 주세요.');
    return res.json();
  });
}

function loadIndex_(){
  if (indexCache) return Promise.resolve(indexCache);
  return fetchJson_('standards-index.json').then(function(d){ indexCache = d; return d; });
}

/** 전략은행을 부팅 때 미리 받아 두어 CBIL 원문(buildBankText_)의 동기 호출을 살린다 */
function preloadBank_(){
  if (bankCache) return Promise.resolve(bankCache);
  return fetchJson_('strategies.json').then(function(d){
    var rows = [];
    var src = (d && d.strategies) || [];
    for (var i = 0; i < src.length; i++) {
      rows.push({ phase: src[i].stage || '', sub: src[i].type || '', name: src[i].name || '',
                  desc: src[i].desc || '', page: src[i].page || '', src: src[i].src || '' });
    }
    bankCache = { rows: rows };
    return bankCache;
  })['catch'](function(){ bankCache = { rows: [] }; return bankCache; }); // 실패해도 앱은 계속(내장 폴백)
}

/** CBIL.gs 원문이 부르는 동기 함수 — 미리 받아 둔 캐시를 돌려준다 */
function getStrategyBank(){
  if (!bankCache) throw new Error('전략은행이 아직 로드되지 않았습니다.');
  return bankCache;
}

/** [코드] 접두어 제거(Standards.gs stripCode_와 동일) */
function stripCode_(text){
  return String(text == null ? '' : text).trim().replace(/^\[[^\]]*\]\s*/, '');
}

/** data/*.json 항목 → 레거시 행 모양(클라이언트·CBIL이 아는 형태)으로 변환 */
function rowsOfBand_(band){
  if (bandCache[band]) return Promise.resolve(bandCache[band]);
  return loadIndex_().then(function(idx){
    var meta = null;
    for (var i = 0; i < idx.bands.length; i++) if (idx.bands[i].band === band) { meta = idx.bands[i]; break; }
    if (!meta) throw new Error('학년군 "' + band + '" 데이터를 찾을 수 없습니다.');
    return fetchJson_(meta.file).then(function(d){
      var rows = [], src = d.standards || [];
      for (var j = 0; j < src.length; j++) {
        var s = src[j], lv = s.levels || {};
        rows.push({
          gradeBand: s.band, subject: s.subject,
          area: s.domain || '',
          code: s.code || '',
          statement: stripCode_(s.text),
          levelA: lv.A || '', levelB: lv.B || '', levelC: lv.C || '',
          levelD: lv.D || '', levelE: lv.E || '',
          subjectDetail: s.course || ''
        });
      }
      bandCache[band] = rows;
      return rows;
    });
  });
}

function getInitData(){
  return Promise.all([loadIndex_(), preloadBank_()]).then(function(rs){
    var idx = rs[0];
    var bands = [], byBand = {};
    for (var i = 0; i < idx.bands.length; i++) {
      var b = idx.bands[i];
      bands.push(b.band);
      var subs = [];
      for (var j = 0; j < (b.subjects || []).length; j++) subs.push(b.subjects[j].name);
      byBand[b.band] = subs;
    }
    return { gradeBands: bands, subjectsByBand: byBand };
  });
}

function getStandards(gradeBand, subject, area){
  var gb = String(gradeBand || '').trim(), sub = String(subject || '').trim(), ar = area ? String(area).trim() : '';
  return rowsOfBand_(gb).then(function(rows){
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.subject !== sub) continue;
      if (ar && r.area !== ar) continue;
      out.push(r);
    }
    return out;
  });
}

function getAreas(gradeBand, subject){
  return getStandards(gradeBand, subject).then(function(rows){
    var out = [], seen = {};
    for (var i = 0; i < rows.length; i++) {
      var a = rows[i].area;
      if (a && !seen[a]) { seen[a] = true; out.push(a); }
    }
    return out;
  });
}

function getTonghapData(){
  if (tonghapCache) return Promise.resolve(tonghapCache);
  return fetchJson_('tonghap.json').then(function(d){ tonghapCache = d || { units: [] }; return tonghapCache; });
}

/* ══════════ 2. Gemini 계층 (UrlFetchApp 대체 — 사다리 로직 유지) ══════════ */
function getApiKeyOrThrow_(){
  var key = null;
  try { key = window.localStorage.getItem(KEY_STORE); } catch (e) {}
  if (!key) throw new Error('API 키가 등록되어 있지 않아요. 화면 오른쪽 위 🔑 API 키 버튼에서 본인 키를 등록해 주세요(무료, 1분 소요).');
  return key;
}

function hasUserApiKey(){
  try { return !!window.localStorage.getItem(KEY_STORE); } catch (e) { return false; }
}

function saveUserApiKey(key){
  key = String(key || '').trim();
  if (!key) return Promise.reject(new Error('API 키를 붙여넣어 주세요.'));
  if (key.length < 20) return Promise.reject(new Error('키가 너무 짧아요. AI Studio에서 복사한 키 전체를 붙여넣어 주세요.'));
  return testKey_(key).then(function(status){
    if (status === 'invalid') {
      throw new Error('키가 유효하지 않아요. Google AI Studio(aistudio.google.com/app/apikey)에서 키를 다시 복사해 주세요.');
    }
    try { window.localStorage.setItem(KEY_STORE, key); }
    catch (e) { throw new Error('브라우저 저장소에 키를 저장하지 못했습니다. 시크릿 창이 아닌 일반 창에서 시도해 주세요.'); }
    return { ok: true, note: status === 'quota' ? '다만 지금은 호출 한도 상태라 잠시 후부터 사용할 수 있어요.' : '' };
  });
}

function clearUserApiKey(){
  try { window.localStorage.removeItem(KEY_STORE); } catch (e) {}
  return Promise.resolve({ ok: true });
}

/** v3.0: 계정 개념이 없다 — 키 저장 위치를 알려 준다(🔑 창 표시용) */
function whoAmI(){
  return Promise.resolve('이 브라우저');
}

function sleep_(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }

function testKey_(key){
  var url = 'https://generativelanguage.googleapis.com/' + GEMINI_API_VERSION +
            '/models/' + GEMINI_FALLBACK_MODEL + ':generateContent?key=' + encodeURIComponent(key);
  var payload = { contents: [{ parts: [{ text: 'ping' }] }], generationConfig: { maxOutputTokens: 1 } };
  return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) })
    .then(function(res){
      if (res.status === 200) return 'ok';
      if (res.status === 429 || res.status === 503) return 'quota';
      return 'invalid';
    }, function(){ return 'invalid'; });
}

function msgOf_(e){ return (e && e.message) ? e.message : String(e || ''); }
function isQuotaError_(e){ var m = msgOf_(e); return m.indexOf('QUOTA_DAILY') === 0 || m.indexOf('QUOTA_MINUTE') === 0 || m.indexOf('HTTP 429') !== -1; }
function isOverloadError_(e){ var m = msgOf_(e); return m.indexOf('OVERLOADED') === 0 || m.indexOf('HTTP 503') !== -1; }
function isDailyQuotaBody_(body){ return /per\s*day|perday|daily/i.test(String(body || '')); }

function callGemini(prompt, schema){
  // Gemini가 이따금 형식이 어긋난 JSON을 반환한다(간헐적). 복구 실패 시 같은 요청을 1회만 자동 재시도한다.
  return callGeminiLadder_(prompt, schema)['catch'](function(e){
    if (isParseError_(e)) return callGeminiLadder_(prompt, schema);
    throw e;
  });
}

function isParseError_(e){ return msgOf_(e).indexOf('JSON 파싱 실패') === 0; }

function callGeminiLadder_(prompt, schema){
  return callGeminiModel_(GEMINI_MODEL, prompt, schema)['catch'](function(e){
    if (!isQuotaError_(e) && !isOverloadError_(e)) throw e;
    return callGeminiModel_(GEMINI_FALLBACK_MODEL, prompt, schema)['catch'](function(e2){
      if (isQuotaError_(e2)) {
        var both = msgOf_(e) + ' ' + msgOf_(e2);
        if (both.indexOf('QUOTA_MINUTE') !== -1) {
          throw new Error('요청이 잠깐 몰렸어요(분당 호출 한도). 1분쯤 뒤에 같은 버튼을 다시 눌러 주세요 — 진행 내용은 자동 저장되어 있습니다.');
        }
        throw new Error('오늘 치 무료 호출 한도를 모두 사용했어요. 무료 한도는 한국 시간 오후 4시쯤 다시 채워집니다. 진행 내용은 자동 저장되어 있으니 그때 이어서 하면 됩니다.');
      }
      if (isOverloadError_(e2)) {
        throw new Error('Gemini 서버가 일시적으로 혼잡해요. 잠시 뒤 같은 버튼을 다시 눌러 주세요 — 진행 내용은 자동 저장되어 있습니다.');
      }
      throw e2;
    });
  });
}

function fetchGemini_(url, payload){
  return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) })
    .then(function(res){ return res.text().then(function(body){ return { code: res.status, body: body }; }); },
          function(){ throw new Error('네트워크 오류로 Gemini에 연결하지 못했습니다. 인터넷 연결을 확인해 주세요.'); });
}

function callGeminiModel_(model, prompt, schema){
  var apiKey;
  try { apiKey = getApiKeyOrThrow_(); } catch (e) { return Promise.reject(e); }

  var url = 'https://generativelanguage.googleapis.com/' + GEMINI_API_VERSION +
            '/models/' + model + ':generateContent?key=' + encodeURIComponent(apiKey);
  var generationConfig = { responseMimeType: 'application/json', temperature: GEMINI_TEMPERATURE, maxOutputTokens: GEMINI_MAX_TOKENS };
  if (schema) generationConfig.responseSchema = schema;
  var payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: generationConfig };

  return fetchGemini_(url, payload).then(function(r){
    // 순간 몰림(분당 429·일시 혼잡 503)은 5초 쉬고 한 번 자동 재시도
    if ((r.code === 429 && !isDailyQuotaBody_(r.body)) || r.code === 503) {
      return sleep_(5000).then(function(){ return fetchGemini_(url, payload); });
    }
    return r;
  }).then(function(r){
    if (r.code === 429) throw new Error((isDailyQuotaBody_(r.body) ? 'QUOTA_DAILY' : 'QUOTA_MINUTE') + ' (' + model + ')');
    if (r.code === 503) throw new Error('OVERLOADED (' + model + ')');
    if (r.code !== 200) throw new Error('Gemini API 오류 (HTTP ' + r.code + '): ' + String(r.body || '').substring(0, 800));
    var data = JSON.parse(r.body);
    var text = extractText_(data);
    if (!text) {
      var reason = '';
      if (data.candidates && data.candidates[0] && data.candidates[0].finishReason) reason = ' (finishReason: ' + data.candidates[0].finishReason + ')';
      else if (data.promptFeedback && data.promptFeedback.blockReason) reason = ' (blockReason: ' + data.promptFeedback.blockReason + ')';
      throw new Error('Gemini 응답에 텍스트가 없습니다' + reason + '. 원문: ' + String(r.body || '').substring(0, 500));
    }
    try { return safeParseJson_(text); }
    catch (pe) {
      var repaired = repairJson_(text);
      if (repaired !== null) return repaired; // 잘린 꼬리를 정리해 완성된 후보까지 살림
      var fr = (data.candidates && data.candidates[0] && data.candidates[0].finishReason) ? data.candidates[0].finishReason : '';
      if (fr && fr !== 'STOP') throw new Error(msgOf_(pe) + '\n(응답 중단 사유: ' + fr + ')');
      throw pe;
    }
  });
}

function extractText_(data){
  if (!data || !data.candidates || !data.candidates.length) return '';
  var cand = data.candidates[0];
  if (!cand.content || !cand.content.parts || !cand.content.parts.length) return '';
  var parts = cand.content.parts, text = '';
  for (var i = 0; i < parts.length; i++) if (parts[i].text) text += parts[i].text;
  return text;
}

function safeParseJson_(text){
  var t = String(text).trim();
  t = t.replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();
  try { return JSON.parse(t); }
  catch (e) { throw new Error('JSON 파싱 실패: ' + e + '\n원문(앞 500자): ' + t.substring(0, 500)); }
}

/**
 * 끝이 잘리거나 형식이 어긋난 JSON에서, 온전한 부분까지만 살려 파싱을 시도한다.
 * 뒤에서부터 '값의 끝일 만한 지점'(닫는 괄호·따옴표)으로 잘라 가며 열린 괄호를 닫아 본다.
 * 성공하면 객체를, 전부 실패하면 null을 반환(호출부가 자동 재시도로 넘어감).
 */
function repairJson_(text){
  var t = String(text).trim().replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();
  var tries = 0;
  for (var cut = t.length; cut > 1 && tries < 400; cut--) {
    var ch = t.charAt(cut - 1);
    if (ch !== '}' && ch !== ']') continue; // 완결된 객체·배열 경계에서만 자름(반쪽 항목 방지)
    tries++;
    var head = t.substring(0, cut).replace(/,\s*$/, '');
    var closed = closeBrackets_(head);
    if (closed === null) continue;
    try {
      var obj = JSON.parse(closed);
      // 살려낸 분량이 원문의 60% 미만이면 버림 — 이럴 땐 자동 재시도가 더 낫다
      if (JSON.stringify(obj).length < t.length * 0.6) return null;
      return obj;
    } catch (e) {}
  }
  return null;
}

/** 문자열 상태를 추적하며 열린 { [ 를 세어, 부족한 닫는 괄호를 붙인다. 구조가 어긋나면 null. */
function closeBrackets_(s){
  var stack = [], inStr = false, esc = false;
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    if (inStr) {
      if (esc) esc = false;
      else if (c === '\\') esc = true;
      else if (c === '"') inStr = false;
      continue;
    }
    if (c === '"') inStr = true;
    else if (c === '{' || c === '[') stack.push(c);
    else if (c === '}') { if (stack.pop() !== '{') return null; }
    else if (c === ']') { if (stack.pop() !== '[') return null; }
  }
  if (inStr) return null; // 문자열 한복판에서 끊긴 지점 — 이 절단점은 포기
  var tail = '';
  for (var j = stack.length - 1; j >= 0; j--) tail += (stack[j] === '{' ? '}' : ']');
  return s + tail;
}

/* ══════════ 3. CBIL 단계 로직 — v2.5 CBIL.gs 원문 (무수정) ══════════ */
/**
 * CBIL 단원설계 앱 — 모듈 ④ CBIL 단계 로직 (CBIL.gs)
 * ------------------------------------------------------------------
 * CBIL 1단계(개념 추출)의 미니스텝별 프롬프트를 둔다. 각 함수는 JSON만 반환한다.
 * 클라이언트가 맥락(ctx)을 넘기면, 여기서 프롬프트를 만들어 callGemini를 호출한다.
 * (callGemini는 Gemini.gs에 있음 — JSON 모드 + 파싱 처리)
 *
 * 거대한 단일 프롬프트를 피하고, 미니스텝마다 작은 프롬프트를 쓴다.
 */

/** 선택 성취기준을 프롬프트용 텍스트로 변환 */
function standardsToText_(standards) {
  var lines = [];
  for (var i = 0; i < standards.length; i++) {
    lines.push('- ' + standards[i].code + ' ' + standards[i].statement);
  }
  return lines.join('\n');
}

/**
 * 학년군 문자열로 학교급을 가린다. 반환: 'elemLow' | 'elemHigh' | 'middle' | 'high'
 * 중·고는 학년군에 "중학교"/"고등학교"가 들어오고, 초등은 코드(12 / 34 / 5~6)로 저/중고학년 구분.
 */
function gradeStage_(gradeBand) {
  var g = (gradeBand || '').toString();
  if (g.indexOf('고등') > -1 || g.indexOf('고교') > -1) return 'high';
  if (g.indexOf('중학') > -1) return 'middle';
  if (g.indexOf('12') > -1) return 'elemLow';
  if (g.indexOf('1') > -1 && g.indexOf('2') > -1 && g.indexOf('3') === -1 && g.indexOf('4') === -1) return 'elemLow';
  return 'elemHigh';
}

/** 페르소나·문맥용 학교급 표기. '초등' | '중등' | '고등' */
function gradeStageText_(gradeBand) {
  var s = gradeStage_(gradeBand);
  if (s === 'high') return '고등';
  if (s === 'middle') return '중등';
  return '초등';
}

/** 학년 수준에 맞는 학생 사고·표현 톤(명사 앞에 붙는 형용사구). */
function gradeToneText_(gradeBand) {
  var s = gradeStage_(gradeBand);
  if (s === 'high') return '고등학생 눈높이에서, 추상적·비판적·다각적으로 사고하도록 학문적 개념어와 정교한 표현을 살려 단순화하지 않은';
  if (s === 'middle') return '중학생 눈높이에서, 추상 개념을 구체적 사례로 비계하고 분석적으로 사고하도록 학문 어휘를 풀어 쓴';
  if (s === 'elemLow') return '초등 저학년 눈높이에서, 아주 구체적이고 놀이처럼 감각·경험에 닿는 쉬운';
  return '초등 중·고학년 눈높이에서, 구체적 사례와 일상에 연결한 친근하고 쉬운';
}

/**
 * 1-ⓐ 개념적 렌즈 후보 생성.
 * @param {Object} ctx { gradeBand, subject, standards:[{code, statement}, ...] }
 * @return {Object} { lenses: [{ name, rationale }, ...] }
 */
function generateLenses(ctx) {
  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '아래는 한 교사가 이번 단원에서 다룰 성취기준이야.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '이 성취기준들을 관통할 수 있는 "개념적 렌즈(conceptual lens)" 후보를 4개 제안해줘.\n' +
    '- 개념적 렌즈는 단원 전체를 꿰뚫는 상위의 전이 가능한 개념이야 ' +
    '(예: 변화, 관계, 상호작용, 관점, 시스템, 인과, 지속과 변화, 다양성, 질서 등).\n' +
    '- 너무 좁은 소재(예: 옛날 물건)가 아니라, 여러 사실을 묶어 일반화를 이끌어낼 수 있는 추상적 개념이어야 해.\n' +
    '- 서로 다른 관점의 렌즈를 섞어서 제안해(교사가 고를 수 있도록).\n' +
    '각 렌즈마다, 왜 이 성취기준들에 적합한지 교사가 이해하기 쉬운 말로 1~2문장 근거를 달아줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"lenses":[{"name":"렌즈 이름","rationale":"이 성취기준에 적합한 이유"}]}';

  return callGemini(prompt);
}

/**
 * 1-ⓑ 단원명 후보 생성.
 * @param {Object} ctx { gradeBand, subject, standards:[{code, statement}, ...], lens:{name} }
 * @return {Object} { titles: [{ title, note }, ...] }
 */
function generateTitles(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '아래 정보를 바탕으로, 학생의 호기심을 자극하고 단원의 초점을 분명히 드러내는 "단원명" 후보를 4개 제안해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '좋은 단원명의 조건:\n' +
    '- 생각할 거리를 던지는 질문형도 좋음 (예: "권력은 인간성을 타락시키는가?")\n' +
    '- 구체적인 맥락을 담은 형태도 좋음 (예: "우리는 누구인가: 옛날과 오늘날의 가족")\n' +
    '- 흥미를 끄는 표현도 환영\n' +
    '- 위 개념적 렌즈("' + lensName + '")가 단원명에 자연스럽게 녹아 있어야 함\n' +
    '피해야 할 것: 너무 광범위함(예: 삶, 변화), 너무 좁은 소재 나열, 개념과 무관한 제목, 모호한 제목.\n' +
    '서로 다른 유형(질문형/맥락형 등)을 섞어서 제안하고, ' + gradeToneText_(ctx.gradeBand) + ' 말로 써줘.\n' +
    '각 제목마다, 어떤 점이 좋은지 아주 짧은 한 줄 설명을 달아줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"titles":[{"title":"단원명","note":"이 제목의 장점 한 줄"}]}';

  return callGemini(prompt);
}

/**
 * 1-ⓒ 핵심 아이디어 후보 생성.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title }
 * @return {Object} { coreIdeas: [{ statement, note }, ...] }
 */
function generateCoreIdeas(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var title = ctx.title || '';

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '아래 정보를 바탕으로, 이 단원을 관통하는 "핵심 아이디어(Core Idea)" 후보를 3개 제안해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + title + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '핵심 아이디어는 이 단원에서 학생이 도달하길 바라는 가장 중요한 일반화(전이 가능한 깨달음)야.\n' +
    '- 개념적 렌즈("' + lensName + '")와 성취기준의 핵심 개념을 연결한 완결된 한 문장으로 써줘.\n' +
    '- 시간·문화·상황을 초월해 적용되도록 현재 시제 평서문으로 서술해 — "~합니다/~해요" 같은 높임말이 아니라 "~한다/~된다"로 끝맺어 (예: "사회의 모습은 시간이 흐르며 변화한다").\n' +
    '- 고유명사나 특정 사례에 묶이지 않게, 개념 수준으로 일반화해.\n' +
    '- ' + gradeToneText_(ctx.gradeBand) + ' 말로.\n' +
    '- 서로 다른 관점이나 깊이의 진술을 섞어서 제안해.\n' +
    '각 핵심 아이디어마다, 무엇을 담고 있는지 아주 짧은 한 줄 설명을 달아줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"coreIdeas":[{"statement":"핵심 아이디어 문장","note":"이 진술이 담은 것 한 줄"}]}';

  return callGemini(prompt);
}

/**
 * 1-ⓓ 스트랜드 후보 생성.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea }
 * @return {Object} { strands: [{ name, note }, ...] }
 */
function generateStrands(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '아래 정보를 바탕으로, 이 단원을 다루기 쉬운 부분으로 나누는 "스트랜드(strand)" 후보를 5개 제안해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '스트랜드는 단원을 구성하는 소재나 차원의 묶음이야. 교사는 이 중 3~4개를 골라 단원의 뼈대로 삼을 거야.\n' +
    '- 각 스트랜드는 성취기준의 내용을 실제로 다룰 수 있는 묶음이어야 하고, 후보들을 모으면 성취기준 전체를 빠짐없이 덮을 수 있어야 해.\n' +
    '- 너무 큰 덩어리(단원 전체)나 너무 작은 조각(차시 1개 분량)은 피해.\n' +
    '- 개념적 렌즈와 핵심 아이디어로 자연스럽게 이어질 수 있는 묶음으로.\n' +
    '- 이름은 학생이 이해할 짧은 구로 써줘.\n' +
    '각 스트랜드마다 무엇을 다루는지 한 줄 설명을 달아줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"strands":[{"name":"스트랜드 이름","note":"무엇을 다루는지 한 줄"}]}';

  return callGemini(prompt);
}

/**
 * 1-ⓔ 스트랜드별 관련 개념 생성.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[이름,...] }
 * @return {Object} { byStrand: { "<스트랜드>": ["개념", ...], ... } }
 */
function generateStrandConcepts(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strands = (ctx.strands && ctx.strands.length) ? ctx.strands : [];
  var parts = []; for (var i = 0; i < strands.length; i++) parts.push('"' + strands[i] + '"');
  var strandList = parts.length ? parts.join(', ') : '(미정)';

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '교사가 이미 정한 스트랜드마다, 그 안에서 학생이 학습할 "관련 개념"을 뽑아줘. ' +
    '이 개념들은 다음 단계에서 일반화 문장(두 개 이상의 개념 사이의 관계)을 만드는 재료가 돼.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드 목록: ' + strandList + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '조건:\n' +
    '- 위 스트랜드 하나하나마다 관련 개념을 3~6개씩 제안해.\n' +
    '- "개념"은 사실이나 활동이 아니라 여러 사례에서 추상화되는 전이 가능한 단어·짧은 구여야 해 (예: 변화, 상호의존, 권리, 표현, 균형). 고유명사·특정 사건·차시 활동명은 피해.\n' +
    '- 그 학년 학생이 이해할 수 있는 수준으로, 개념적 렌즈·핵심 아이디어와 자연스럽게 이어지게.\n' +
    '- 키는 위에 준 스트랜드 이름을 그대로 사용해.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해(각 키는 스트랜드 이름, 값은 개념 문자열 배열):\n' +
    '{"byStrand":{"스트랜드이름":["개념1","개념2","개념3"]}}';

  return callGemini(prompt);
}

/**
 * 1-ⓕ 예상 선개념·오개념 후보 생성.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[이름,...] }
 * @return {Object} { preconceptions: [{ statement, note }, ...] }
 */
function generatePreconceptions(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strandText = (ctx.strands && ctx.strands.length) ? ctx.strands.join(', ') : '(미정)';

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL)과 학습과학에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 전문가야.\n' +
    '『How People Learn』의 첫째 원리처럼, 학생은 백지가 아니라 이미 가진 생각(선개념) 위에서 새 지식을 구성해. ' +
    '그 생각이 오개념이라면 수업이 그것을 끌어내어 정면으로 다뤄야 해.\n' +
    '아래 단원에 대해, 이 학년 학생들이 흔히 가지고 있을 "예상 선개념·오개념" 후보를 5개 제안해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드: ' + strandText + '\n' +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '조건:\n' +
    '- 각 항목은 학생의 머릿속 생각을 한 문장으로 서술해 (예: "옛날 물건은 지금 우리 생활과 아무 상관이 없는 낡은 것이라고 생각한다").\n' +
    '- 진짜 그 학년 아이들이 가질 법한, 구체적이고 현실적인 생각으로 써줘 (어른의 오개념 말고).\n' +
    '- 단원의 핵심 개념 이해를 가로막는 오개념 위주로 하되, 수업의 출발점으로 살릴 수 있는 불완전한 사전지식도 1개쯤 섞어도 좋아.\n' +
    '각 항목마다, 이것이 왜 이 단원에서 중요한지(어떤 개념 이해와 부딪히는지, 또는 어떻게 활용할 수 있는지) 한 줄 설명을 달아줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"preconceptions":[{"statement":"학생의 생각 한 문장","note":"왜 중요한지 한 줄"}]}';

  return callGemini(prompt);
}

/**
 * 2단계 일반화 진술문 후보 생성 (스트랜드별 + 렌즈 차원).
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[...], strandConcepts:{}, preconceptions:[...] }
 * @return {Object} { byStrand: { "<스트랜드>": [{statement,note}] }, lens: [{statement,note}] }
 */
function generateGeneralizations(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strands = (ctx.strands && ctx.strands.length) ? ctx.strands : [];
  var sc = ctx.strandConcepts || {};
  var strandLines = [];
  for (var i = 0; i < strands.length; i++) {
    var nm = strands[i];
    var cs = (sc[nm] && sc[nm].length) ? sc[nm].join(', ') : '(관련 개념 미정)';
    strandLines.push('- "' + nm + '" — 관련 개념: ' + cs);
  }
  var preconText = '';
  if (ctx.preconceptions && ctx.preconceptions.length) {
    var pl = [];
    for (var p = 0; p < ctx.preconceptions.length; p++) pl.push('- ' + ctx.preconceptions[p]);
    preconText = '예상 선개념·오개념:\n' + pl.join('\n') + '\n';
  }

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '학생들이 이 단원의 탐구 끝에 스스로 구성하게 될 "일반화(generalization)" 문장 후보를, ' +
    '아래 스트랜드마다 3개씩 + 개념적 렌즈 차원에서 2개 제안해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드와 관련 개념:\n' + (strandLines.length ? strandLines.join('\n') : '(미정)') + '\n' +
    preconText +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '일반화 작성 조건 (Erickson의 CBCI 원리):\n' +
    '- 각 스트랜드의 일반화는 **그 스트랜드의 관련 개념 가운데 두 개 이상**의 관계를 서술해야 해. 그 개념 단어가 문장 안에 실제로 보이게 써줘.\n' +
    '- 렌즈 차원 일반화 2개는 특정 스트랜드를 넘어 단원 전체를 관통하는 이해로, 렌즈("' + lensName + '")의 시선이 드러나야 하고 핵심 아이디어와 같은 방향이어야 해.\n' +
    '- 모든 문장은 시간·문화·상황을 초월해 적용되도록 현재 시제 평서문으로("~합니다/~해요" 높임말이 아니라 "~한다/~된다"로 끝맺기), 고유명사나 특정 사례 없이.\n' +
    '- 약한 동사("~와 관련이 있다", "~에 영향을 준다")보다 관계를 구체적으로 드러내는 동사를 써줘.\n' +
    '- 위 예상 오개념을 바로잡는 방향의 일반화가 1~2개 포함되면 좋아.\n' +
    '- 학생이 탐구 끝에 자기 말로 말할 수 있을, ' + gradeToneText_(ctx.gradeBand) + ' 문장으로.\n' +
    '각 문장의 note에는 어떤 개념들을 연결하는지(오개념 교정이면 그것도) 한 줄로 적어줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해(byStrand의 키는 위 스트랜드 이름을 그대로):\n' +
    '{"byStrand":{"스트랜드이름":[{"statement":"일반화 문장","note":"연결 개념 한 줄"}]},"lens":[{"statement":"일반화 문장","note":"연결 개념 한 줄"}]}';

  return callGemini(prompt);
}

/**
 * 3단계 탐구 질문 세트 생성 (사실적·개념적·논쟁적·메타인지).
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[...], preconceptions:[...], generalizations:[...] }
 * @return {Object} { questions: { factual:[], conceptual:[], debatable:[], metacognitive:[] } }
 */
function generateQuestions(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strandText = (ctx.strands && ctx.strands.length) ? ctx.strands.join(', ') : '(미정)';

  var gens = ctx.generalizationsDetailed || [];
  var sc = ctx.strandConcepts || {};
  var strandsArr2 = (ctx.strands && ctx.strands.length) ? ctx.strands : [];
  var scLines = [];
  for (var sci = 0; sci < strandsArr2.length; sci++) {
    var snm = strandsArr2[sci];
    var scs = (sc[snm] && sc[snm].length) ? sc[snm].join(', ') : '(미정)';
    scLines.push('- "' + snm + '": ' + scs);
  }
  var scText = '스트랜드별 관련 개념(개념 형성 질문이 다룰 개념):\n' + (scLines.length ? scLines.join('\n') : '(미정)') + '\n';
  if (!gens.length && ctx.generalizations) {
    for (var f = 0; f < ctx.generalizations.length; f++) gens.push({ statement: ctx.generalizations[f], strand: '' });
  }
  var genLines = [];
  for (var i = 0; i < gens.length; i++) {
    genLines.push((i + 1) + '. ' + (gens[i].strand ? '[' + gens[i].strand + '] ' : '') + gens[i].statement);
  }
  var preconText = '';
  if (ctx.preconceptions && ctx.preconceptions.length) {
    var pl = [];
    for (var j = 0; j < ctx.preconceptions.length; j++) pl.push('- ' + ctx.preconceptions[j]);
    preconText = '예상 선개념·오개념:\n' + pl.join('\n') + '\n';
  }

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL) 단원 설계를 돕는 전문가야.\n' +
    '아래 일반화 문장 하나하나에 대해 "질문 세트"를 만들어줘. 질문은 학생을 사실에서 개념으로, 개념에서 그 일반화로 안내하는 길이야.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드: ' + strandText + '\n' +
    '일반화 목록(번호 기준):\n' + (genLines.length ? genLines.join('\n') : '(미정)') + '\n' +
    scText +
    preconText +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '만들 것:\n' +
    '1) 위 일반화 각각에 대해 — (가) 사실적 질문 3개: 그 일반화의 토대가 될 사실·내용을 성취기준 범위에서 묻는, 답이 정해진 질문. (나) 개념 형성 질문 2~3개: 그 일반화가 속한 스트랜드의 관련 개념 하나하나(정의·특징·속성)를 학생이 스스로 세워 가도록 돕는 질문 — 개념 사이의 관계가 아니라 개념 하나하나의 뜻에 초점을 두고, 관련 개념마다 하나씩 만들어. 단 개념 형성 질문은 다음 세 규칙을 반드시 지켜: ① 개념어(예: 상호의존·교류·협력)의 뜻을 질문 안에서 풀어 설명하지 마 — 개념어 앞에 그 정의를 수식어로 붙이면 답이 새어 나가 학생이 생각할 게 없어진다. ② 대신 학생이 같은 일반화의 사실적 질문에서 살펴본 구체적 사례·경험을 떠올려, 그 공통점을 찾거나 한 낱말로 이름 붙이도록 이끄는 귀납형으로 써. ③ 개념어는 질문이 향하는 도착점이지 질문 안에 미리 박아 두는 전제가 아니야 — 그러니 개념 형성 질문에는 개념어를 굳이 넣지 않아도 된다(앞의 핵심 개념 단어를 보이게 하라는 지시는 사실적·개념적 질문에만 적용). 예시 — 나쁜 예 "서로 도움을 주고받는 상호의존이란 어떤 뜻일까요?"는 정의가 이미 들어 있어 생각할 여지가 없으니 이렇게 쓰지 마. 좋은 예 "여러 지역이 부족한 것을 서로 주고받는 이런 모습을, 우리는 한 낱말로 뭐라고 부르면 좋을까요?"처럼 사례에서 개념을 끌어내게 해. (다) 개념적 질문 3개: 그 일반화 속 개념들의 관계를 묻는, 시간·상황을 넘어 적용되는 질문(고유명사 없이). 질문에 그 일반화/스트랜드의 핵심 개념 단어가 보이게 해줘.\n' +
    '2) 단원 전체에 대해 — 논쟁적 질문 3개(정답이 하나가 아니어서 입장과 근거를 요구하는 질문, 위 오개념과 부딪히는 지점이면 좋아) + 메타인지 질문 3개(학생이 자기 생각과 배움을 점검하는 질문).\n' +
    '모두 ' + gradeToneText_(ctx.gradeBand) + ' 의문문으로, 물음표로 끝나게 써줘.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해(index는 위 일반화 목록의 번호):\n' +
    '{"byGen":[{"index":1,"factual":["질문"],"conceptForm":["질문"],"conceptual":["질문"]}],"debatable":["질문"],"metacognitive":["질문"]}';

  return callGemini(prompt);
}

/**
 * 4단계 최종 수행 평가 생성 — GRASPS 시나리오 3개 + 형성평가 체크포인트.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[...], preconceptions:[...], generalizations:[...] }
 * @return {Object} { scenarios:[{title,pitch,goal,role,audience,situation,product,standard,transferNote}], checkpoints:[{name,note}] }
 */
function generateGrasps(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strandText = (ctx.strands && ctx.strands.length) ? ctx.strands.join(', ') : '(미정)';
  var genText = '';
  if (ctx.generalizations && ctx.generalizations.length) {
    var gl = [];
    for (var i = 0; i < ctx.generalizations.length; i++) gl.push('- ' + ctx.generalizations[i]);
    genText = '일반화(평가가 겨냥할 이해):\n' + gl.join('\n') + '\n';
  }

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL)과 백워드 설계(Wiggins & McTighe)에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 전문가야.\n' +
    '아래 단원의 "최종 수행 평가"를 GRASPS 구조로 설계해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드: ' + strandText + '\n' +
    genText +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '요청 1) 서로 성격이 다른 GRASPS 수행 과제 시나리오를 3개 제안해줘. 각 시나리오 조건:\n' +
    '- 학생이 위 일반화를 *수업에서 직접 다루지 않은 새로운 실제적 맥락*에 적용하는 전이 과제여야 해(원전이).\n' +
    '- goal(과제가 요구하는 목표), role(학생이 맡는 역할), audience(결과물을 받아볼 실제적 청중), situation(과제가 놓인 실제적 상황), product(만들어낼 결과물과 핵심 요건), standard(평가 기준 — 지식/기능/태도 각 1줄씩, "지식: …\\n기능: …\\n태도: …" 형태)를 모두 채워줘.\n' +
    '- title(과제 제목)과 pitch(교사가 한눈에 파악할 한 줄 소개), transferNote(어떤 점에서 새로운 맥락으로의 전이인지 한 줄)도 함께.\n' +
    '- ' + gradeStageText_(ctx.gradeBand) + ' 교실에서 준비물·시간 면에서 실제로 실행 가능한 과제로.\n\n' +
    '요청 2) 단원 곳곳에 배치할 "형성평가 체크포인트"를 5개 제안해줘. 조건:\n' +
    '- 점수 부담이 없는 저부담 인출 활동(예: 3분 쪽지 퀴즈, 브레인 덤프, 출구 티켓, 짝 설명하기, 개념 스케치).\n' +
    '- name에는 활동 이름, note에는 무엇을 떠올리게 하는지와 추천 시점(예: 개념연결 단계 후)을 한 줄로.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"scenarios":[{"title":"","pitch":"","goal":"","role":"","audience":"","situation":"","product":"","standard":"","transferNote":""}],"checkpoints":[{"name":"","note":""}]}';

  return callGemini(prompt);
}

/**
 * 4단계 최종 수행 평가 (RAFTS) — 맥락 있는 쓰기·수행 과제 시나리오 3개 + 채점 루브릭 + 형성평가 체크포인트.
 * RAFTS = Role(역할)·Audience(독자)·Format(형식)·Topic(주제)·Strong Verb(핵심 동사).
 * generateGrasps와 동일한 단원 맥락(ctx)을 받아 단일 호출로 생성한다(추가 API 비용 없음).
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands:[...], generalizations:[...] }
 * @return {Object} { scenarios:[{title,pitch,role,audience,format,topic,strongVerb,task,transferNote}], rubric:[{criterion,A,B,C}], checkpoints:[{name,note}] }
 */
function generateRafts(ctx) {
  var RUBRIC_AXES = ['내용', '조직', '표현']; // ← 지식·기능·태도로 바꾸려면 이 한 줄만 교체

  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var strandText = (ctx.strands && ctx.strands.length) ? ctx.strands.join(', ') : '(미정)';
  var genText = '';
  if (ctx.generalizations && ctx.generalizations.length) {
    var gl = [];
    for (var i = 0; i < ctx.generalizations.length; i++) gl.push('- ' + ctx.generalizations[i]);
    genText = '일반화(평가가 겨냥할 이해):\n' + gl.join('\n') + '\n';
  }
  var axisText = RUBRIC_AXES.join(' · ');
  var rubricShape = [];
  for (var a = 0; a < RUBRIC_AXES.length; a++) {
    rubricShape.push('    {"criterion":"' + RUBRIC_AXES[a] + '","A":"","B":"","C":""}');
  }

  var prompt =
    '너는 2022 개정 교육과정의 개념 기반 탐구 학습(CBIL)과 백워드 설계(Wiggins & McTighe)에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 전문가야.\n' +
    '아래 단원의 "최종 수행 평가"를 RAFTS 구조로 설계해줘. RAFTS는 맥락 있는 쓰기·수행 과제를 위한 틀이야: Role(역할)·Audience(독자)·Format(형식)·Topic(주제)·Strong Verb(핵심 동사).\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드: ' + strandText + '\n' +
    genText +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '요청 1) 서로 성격이 다른 RAFTS 수행 과제 시나리오를 3개 제안해줘. 각 시나리오 조건:\n' +
    '- 학생이 위 일반화를 *수업에서 직접 다루지 않은 새로운 실제적 맥락*에 적용하는 전이 과제여야 해(원전이).\n' +
    '- strongVerb(핵심 동사)를 먼저 정한다는 느낌으로, 분석·평가·설득·비교처럼 고차원 사고(블룸 분류의 분석·평가·창조)를 요구하는 동사를 골라줘.\n' +
    '- role(학생이 맡는 역할), audience(글/결과물을 받아볼 실제적 독자), format(글/결과물의 형식 — 예: 편지·기사·제안서·안내문·일기), topic(다룰 핵심 주제), strongVerb(핵심 동사)를 모두 채워줘. role·audience·format이 현실적으로 일관되게.\n' +
    '- title(과제 제목), pitch(교사가 한눈에 파악할 한 줄 소개), transferNote(어떤 점에서 새로운 맥락으로의 전이인지 한 줄)도 함께.\n' +
    '- task(과제 지시문): 위 R·A·F·T·S를 한 단락으로 엮어 학생에게 직접 주는 지시문으로. 학생을 향한 명령·청유형으로 써(예: "~을 작성하라", "~해 보자"). ' + gradeStageText_(ctx.gradeBand) + ' 교실에서 실제로 실행 가능하게.\n\n' +
    '요청 2) 위 과제를 채점할 루브릭을 1개 만들어줘. 조건:\n' +
    '- 평가 축 ' + RUBRIC_AXES.length + '개: ' + axisText + '.\n' +
    '- 각 축마다 성취수준 A(상)·B(중)·C(하)를 학생 수행으로 관찰 가능하게 진술해. "~한다/~된다" 평서문으로, 높임말·미사여구 금지.\n' +
    '- 단편 사실 확인이 아니라, 일반화의 관계(왜·어떻게·만약에)를 과제 수행에서 드러내는지를 기준으로 삼아.\n\n' +
    '요청 3) 단원 곳곳에 배치할 "형성평가 체크포인트"를 5개 제안해줘. 점수 부담이 없는 저부담 인출 활동(예: 3분 쪽지 퀴즈, 브레인 덤프, 출구 티켓, 짝 설명하기). name에 활동 이름, note에 무엇을 떠올리게 하는지와 추천 시점을 한 줄로.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"scenarios":[{"title":"","pitch":"","role":"","audience":"","format":"","topic":"","strongVerb":"","task":"","transferNote":""}],"rubric":[\n' +
    rubricShape.join(',\n') + '\n],"checkpoints":[{"name":"","note":""}]}';

  return callGemini(prompt);
}

/**
 * 5단계 탐구 수업 흐름 생성 — 마샬·프렌치(2018) 개념 기반 탐구 7단계별 활동 후보.
 * @param {Object} ctx { gradeBand, subject, standards, lens:{name}, title, coreIdea, strands, preconceptions, generalizations, grasps:{title,product}, checkpoints:[...] }
 * @return {Object} { flow: { engage:[{name,note,science}], focus:[...], investigate:[...], organize:[...], generalize:[...], transfer:[...], reflect:[...] } }
 */
function generateFlowStrand(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var sf = ctx.strandFocus || {};
  var sName = sf.name || '(미정)';
  var isFirst = (sf.order === 1);
  var isLast = (sf.order && sf.total && sf.order === sf.total);
  var fwLabel = (ctx.assessFramework === 'rafts') ? 'RAFTS' : 'GRASPS';

  var preconText = '';
  if (ctx.preconceptions && ctx.preconceptions.length) {
    var pl = [];
    for (var i = 0; i < ctx.preconceptions.length; i++) pl.push('- ' + ctx.preconceptions[i]);
    preconText = '예상 선개념·오개념(이 스트랜드와 관련된 것이 있으면 관계 맺기에서 드러내고, 집중하기에서 정면으로 다룰 것):\n' + pl.join('\n') + '\n';
  }
  var genText = '';
  if (sf.generalizations && sf.generalizations.length) {
    var gl = [];
    for (var j = 0; j < sf.generalizations.length; j++) gl.push('- ' + sf.generalizations[j]);
    genText = '이 스트랜드의 일반화(일반화하기 단계에서 학생들이 스스로 도달해야 할 도착점 — 활동에서 미리 알려주지 말 것):\n' + gl.join('\n') + '\n';
  }
  var lensGenText = '';
  if (sf.lensGeneralizations && sf.lensGeneralizations.length) {
    lensGenText = '단원 전체(렌즈 차원)의 일반화 — 이 사이클이 기여해야 할 큰 그림: ' + sf.lensGeneralizations.join(' / ') + '\n';
  }
  var qText = '';
  if (sf.factualQuestions && sf.factualQuestions.length) {
    qText = '이 스트랜드의 탐구 질문(조사·정리 활동이 이 질문들에 답하게 할 것) — 사실적: ' + sf.factualQuestions.join(' / ')
      + (sf.conceptualQuestions && sf.conceptualQuestions.length ? ' · 개념적: ' + sf.conceptualQuestions.join(' / ') : '') + '\n';
  }
  var graspsText = '';
  if (ctx.grasps && ctx.grasps.title) {
    graspsText = '단원의 최종 수행 과제: "' + ctx.grasps.title + '" — 결과물: ' + (ctx.grasps.product || '')
      + (isLast ? ' ← 마지막 스트랜드이므로 전이하기에 이 과제의 실행을 반드시 포함해.' : ' (이 스트랜드의 전이는 이 과제로 가는 디딤돌이 되게.)') + '\n';
  }
  var cpText = '';
  if (ctx.checkpoints && ctx.checkpoints.length) {
    cpText = '계획된 형성평가 체크포인트(활동 흐름에 자연스럽게 끼워 넣을 것): ' + ctx.checkpoints.join(', ') + '\n';
  }

  var prompt =
    '너는 마샬과 프렌치(Marschall & French, 2018)의 개념 기반 탐구(Concept-Based Inquiry) 모델과 학습과학에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 전문가야.\n' +
    '이 단원은 스트랜드마다 7단계 탐구 사이클(관계 맺기 → 집중하기 → 조사하기 → 조직 및 정리하기 → 일반화하기 → 전이하기 → 성찰하기)을 반복하는 구조야.\n' +
    '지금은 그중 스트랜드 「' + sName + '」' + (sf.order ? ' (' + sf.order + '/' + sf.total + '번째)' : '') + '의 사이클을 설계해줘.\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '이 스트랜드: ' + sName + '\n' +
    '이 스트랜드의 관련 개념(활동이 이 개념들을 다루게 할 것): ' + ((sf.concepts && sf.concepts.length) ? sf.concepts.join(', ') : '(미정)') + '\n' +
    genText + lensGenText + qText + preconText + graspsText + cpText +
    '성취기준:\n' + standardsToText_(ctx.standards) + '\n\n' +
    '각 단계마다 이 스트랜드에 맞는 활동 후보를 2개씩 제안해줘. 단계별 조건:\n' +
    '1) engage(관계 맺기) — 이 스트랜드의 주제로 학생을 지적·감정적으로 끌어들이는 도입. 관련 사전지식을 드러내고 학생이 스스로 질문을 만들게 해.' + (isFirst ? ' 첫 스트랜드이므로 단원 전체의 문을 여는 역할도 겸하게 해줘.' : ' 앞 스트랜드에서 배운 것과 자연스럽게 이어지는 도입이면 좋아.') + '\n' +
    '2) focus(집중하기) — 이 스트랜드의 관련 개념에 대한 공통 이해를 만들어. 예시·비예시 비교, 명확한 정의, 속성 강조로 오개념을 정면으로 다뤄.\n' +
    '3) investigate(조사하기) — 이 스트랜드의 개념과 연결된 사실적 사례·기능을 조사해. 위 사실적 질문에 답하는 조사가 되게, 사례를 바꿔 반복할 수 있는 구조면 좋아.\n' +
    '4) organize(조직 및 정리하기) — 조사한 것을 그래픽 조직자로 정리해 패턴이 드러나게 해(이중부호화, 인지부하 관리).\n' +
    '5) generalize(일반화하기) — 학생들이 패턴에서 위 일반화 문장에 스스로 도달하게 해. 문장 구조 비계("~은 ~ 때문에 ~한다")와 "왜?" 정교화를 써.\n' +
    '6) transfer(전이하기) — 이 스트랜드의 일반화를 새로운 맥락에 적용·검증해. 인출 연습과 적절한 도전을 넣어' + (isLast ? '. 마지막 스트랜드이므로 단원 최종 수행 과제(' + fwLabel + ')의 실행을 포함해.' : ', 최종 수행 과제로 이어지는 디딤돌이 되게 해.') + '\n' +
    '7) reflect(성찰하기) — 이 사이클에서 생각이 어떻게 변했는지 점검(사전·사후 비교, 자기설명). 다음 스트랜드로 넘어가기 전 짧게 할 수 있는 형태로.\n\n' +
    '각 활동: name(짧은 활동 이름), note(교실에서 어떻게 하는지 1~2문장, ' + gradeStageText_(ctx.gradeBand) + ' 수준에서 실행 가능하게), science(기대는 학습과학 원리를 아주 짧게, 예: "선개념 활성화 · HPL 원리1", "인출 연습 · 분산").\n' +
    '형성평가 체크포인트가 자연스럽게 들어갈 활동에는 note에 그 지점을 언급해도 좋아.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"flow":{"engage":[{"name":"","note":"","science":""}],"focus":[{"name":"","note":"","science":""}],"investigate":[{"name":"","note":"","science":""}],"organize":[{"name":"","note":"","science":""}],"generalize":[{"name":"","note":"","science":""}],"transfer":[{"name":"","note":"","science":""}],"reflect":[{"name":"","note":"","science":""}]}}';

  return callGemini(prompt);
}

/**
 * 6단계 활동 전략 생성 — 마샬·프렌치 전략 은행 기반, 단계별 사고 전략 후보.
 * @param {Object} ctx contextForFlow_() + flow:{engage:[활동명...], ...}
 * @return {Object} { strategies: { engage:[{name,note,science}], focus:[...], investigate:[...], organize:[...], generalize:[...], transfer:[...], reflect:[...] } }
 */
function generateStrategies(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';

  var genText = '';
  if (ctx.generalizations && ctx.generalizations.length) {
    var gl = [];
    for (var j = 0; j < ctx.generalizations.length; j++) gl.push('- ' + ctx.generalizations[j]);
    genText = '일반화(전략이 학생을 데려갈 도착점):\n' + gl.join('\n') + '\n';
  }
  var graspsText = '';
  if (ctx.grasps && ctx.grasps.title) {
    graspsText = '최종 수행 과제: "' + ctx.grasps.title + '"\n';
  }

  var PHASE_NAMES = { engage: '관계 맺기', focus: '집중하기', investigate: '조사하기', organize: '조직 및 정리하기', generalize: '일반화하기', transfer: '전이하기', reflect: '성찰하기' };
  var flowText = '';
  if (ctx.flow) {
    var fl = [];
    for (var key in PHASE_NAMES) {
      if (ctx.flow[key] && ctx.flow[key].length) fl.push('- ' + PHASE_NAMES[key] + ': ' + ctx.flow[key].join(' / '));
    }
    if (fl.length) flowText = '단계별로 이미 확정된 활동(전략은 이 활동들과 맞물려야 함):\n' + fl.join('\n') + '\n';
  }

  // [v2.5] 전략 은행 — 시트 「전략은행」 탭에서 읽기 (없으면 내장 폴백 목록)
  var bankText = buildBankText_(PHASE_NAMES);

  var prompt =
    '너는 마샬과 프렌치(Marschall & French, 2018)의 개념 기반 탐구 모델과 학습과학에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 전문가야.\n' +
    '아래 단원의 7단계 수업 흐름에 맞춰, 단계별 "사고 전략"을 제안해줘. 사고 전략은 활동 속에서 학생의 사고를 구조화하는 정형화된 루틴·도구야(활동 자체를 새로 만드는 게 아님).\n\n' +
    '학년군: ' + ctx.gradeBand + '\n' +
    '교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    genText + graspsText + flowText + '\n' +
    bankText +
    '각 단계마다 전략 후보를 3개씩 제안해줘. 조건:\n' +
    '- 은행을 참고하되 단원 맥락에 맞는 변형이나 다른 검증된 전략도 자유롭게 제안해도 좋아. 다만 후보 3개 중 2개 이상은 은행에서 고를 것.\n' +
    '- 같은 단계의 후보 3개는 가능하면 서로 다른 [하위 유형]에서 골라 성격이 겹치지 않게 할 것. 누구나 떠올리는 단골 전략만 반복하지 말고, 이 단원의 렌즈·일반화·활동에 가장 잘 맞물리는 것을 우선할 것.\n' +
    '- name: 전략 이름(은행 전략이면 시트의 이름 그대로 유지).\n' +
    '- note: 이 단원·이 활동에서 구체적으로 어떻게 쓰는지 1~2문장(' + ctx.gradeBand + ' 수준에서 실행 가능, 위에 확정된 활동과 자연스럽게 맞물리게). 은행 전략이고 쪽 번호가 있으면 note 끝에 "(자세한 방법: 책 N쪽)"을 붙일 것.\n' +
    '- science: 기대는 학습과학 원리 아주 짧게(예: "정교화 · 또래 담화", "스키마 조직화 · 인지부하 감소").\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해:\n' +
    '{"strategies":{"engage":[{"name":"","note":"","science":""}],"focus":[{"name":"","note":"","science":""}],"investigate":[{"name":"","note":"","science":""}],"organize":[{"name":"","note":"","science":""}],"generalize":[{"name":"","note":"","science":""}],"transfer":[{"name":"","note":"","science":""}],"reflect":[{"name":"","note":"","science":""}]}}';

  return callGemini(prompt);
}

/**
 * [v2.5] 시트 「전략은행」 탭 → 프롬프트용 텍스트 조립.
 * 탭이 없거나 비어 있으면 짧은 내장 폴백 목록을 반환(앱이 멈추지 않게).
 * 형식: 단계별로 「전략명[하위유형]: 설명 (책 N쪽)」 줄을 만든다.
 */
function buildBankText_(PHASE_NAMES) {
  var bank = null;
  try { bank = getStrategyBank(); } catch (e) { bank = null; }

  if (bank && bank.rows && bank.rows.length) {
    var byPhase = {};
    for (var b = 0; b < bank.rows.length; b++) {
      var it = bank.rows[b];
      if (!byPhase[it.phase]) byPhase[it.phase] = [];
      byPhase[it.phase].push(it);
    }
    var lines = [];
    for (var key in PHASE_NAMES) {
      var pn = PHASE_NAMES[key];
      var list = byPhase[pn] || [];
      if (!list.length) continue;
      var parts = [];
      for (var k = 0; k < list.length; k++) {
        var s = list[k];
        parts.push(s.name +
          (s.sub ? ' [' + s.sub + ']' : '') +
          (s.desc ? ': ' + s.desc : '') +
          (s.page ? ' (책 ' + s.page + '쪽)' : ''));
      }
      lines.push('■ ' + pn + '\n- ' + parts.join('\n- '));
    }
    return '전략 은행(『개념 기반 탐구학습의 실천』 단계별 전략표 기반 · [대괄호]=하위 유형 · 쪽=번역서 쪽 번호):\n' +
           lines.join('\n') + '\n\n';
  }

  // 폴백: 시트 탭이 없을 때만 쓰는 최소 목록 (기존 v2.4 내장 은행 축약)
  return '전략 은행(마샬·프렌치, 축약):\n' +
    '- 관계 맺기: 네 모퉁이 토론, 실험 놀이, 선호도 다이어그램\n' +
    '- 집중하기: 프레이어 모델, 다이아몬드 랭킹, 형용사\n' +
    '- 조사하기: 사례 연구 접근, 실험, 인터뷰·설문\n' +
    '- 조직 및 정리하기: 교차 비교 차트, 시각적 메모, 흐름 다이어그램\n' +
    '- 일반화하기: 스피드 연결, 연결 4, 문장구조(프레임)\n' +
    '- 전이하기: 증명해 봐!, "만약에 ~라면?" 가정 질문, 시사 문제\n' +
    '- 성찰하기: 기준 공동 구성하기, 사전/사후 성찰하기, 체크리스트와 루브릭\n\n';
}

/**
 * 7단계 검토·완성 — 학습과학 자체 점검(10항목 감사) + 학생용 단원 개요 후보.
 * @param {Object} ctx 설계 전체(성취기준~전략, 질문, 전개 순서 포함)
 * @return {Object} { check:[{id,status:'pass'|'warn',evidence,advice}], overviews:[{text,note}] }
 */
function generateReview(ctx) {
  var lensName = (ctx.lens && ctx.lens.name) ? ctx.lens.name : '(미정)';
  var PHASE_NAMES = { engage: '관계 맺기', focus: '집중하기', investigate: '조사하기', organize: '조직 및 정리하기', generalize: '일반화하기', transfer: '전이하기', reflect: '성찰하기' };

  function listText_(label, arr) {
    if (!arr || !arr.length) return '';
    var out = [];
    for (var i = 0; i < arr.length; i++) out.push('- ' + arr[i]);
    return label + ':\n' + out.join('\n') + '\n';
  }
  function phaseMapText_(label, obj) {
    if (!obj) return '';
    var out = [];
    for (var key in PHASE_NAMES) {
      if (obj[key] && obj[key].length) out.push('- ' + PHASE_NAMES[key] + ': ' + obj[key].join(' / '));
    }
    return out.length ? (label + ':\n' + out.join('\n') + '\n') : '';
  }

  var qText = '';
  if (ctx.questions) {
    var q = ctx.questions;
    if (q.byGen && q.byGen.length) {
      var qls = [];
      for (var qi = 0; qi < q.byGen.length; qi++) {
        var b = q.byGen[qi];
        qls.push('- ' + (b.strand ? '[' + b.strand + '] ' : '') + b.gen + ' / 사실: ' + (b.factual || []).join(' · ') + ' / 개념형성: ' + (b.conceptForm || []).join(' · ') + ' / 개념: ' + (b.conceptual || []).join(' · '));
      }
      qText = '일반화별 질문 세트:\n' + qls.join('\n') + '\n'
        + listText_('논쟁적 질문(단원)', q.debatable)
        + listText_('메타인지 질문(단원)', q.metacognitive);
    } else {
      qText = listText_('사실적 질문', q.factual)
        + listText_('개념적 질문', q.conceptual)
        + listText_('논쟁적 질문', q.debatable)
        + listText_('메타인지 질문', q.metacognitive);
    }
  }
  var graspsText = '';
  if (ctx.grasps && ctx.grasps.title) {
    graspsText = '최종 수행 과제(GRASPS): "' + ctx.grasps.title + '" — 결과물: ' + (ctx.grasps.product || '') + '\n'
      + (ctx.grasps.standard ? ('평가 기준: ' + ctx.grasps.standard + '\n') : '');
  }
  var cpText = (ctx.checkpoints && ctx.checkpoints.length) ? ('형성평가 체크포인트: ' + ctx.checkpoints.join(', ') + '\n') : '';
  var seqText = (ctx.sequence && ctx.sequence.length) ? ('단원 전개 순서: ' + ctx.sequence.join(' → ') + '\n') : '';

  var design =
    '학년군: ' + ctx.gradeBand + ' / 교과: ' + ctx.subject + '\n' +
    '단원명: ' + (ctx.title || '') + '\n' +
    '개념적 렌즈: ' + lensName + '\n' +
    '핵심 아이디어: ' + (ctx.coreIdea || '') + '\n' +
    '스트랜드: ' + ((ctx.strands && ctx.strands.length) ? ctx.strands.join(', ') : '') + '\n' +
    listText_('예상 선개념·오개념', ctx.preconceptions) +
    listText_('일반화', ctx.generalizations) +
    qText + graspsText + cpText +
    phaseMapText_('단계별 활동', ctx.flow) +
    phaseMapText_('단계별 사고 전략', ctx.strategies) +
    seqText +
    '성취기준:\n' + standardsToText_(ctx.standards);

  var prompt =
    '너는 개념 기반 탐구 학습(CBIL)과 학습과학(『How People Learn』, 인지과학)에 밝은 ' + gradeStageText_(ctx.gradeBand) + ' 수업 설계 감수자야.\n' +
    '아래는 교사가 완성한 단원 설계 전체야. 두 가지 작업을 해줘.\n\n' +
    '=== 단원 설계 ===\n' + design + '\n=== 설계 끝 ===\n\n' +
    '작업 ① 자체 점검: 아래 10개 항목으로 이 설계를 감사해. 관대하지 말고 정직하게 — 설계 안에 실제 근거가 있을 때만 "pass", 근거가 약하거나 빠졌으면 "warn".\n' +
    '- precon: 선개념·오개념을 끌어내고 정면으로 다루는 활동이 있는가\n' +
    '- frame: 사실이 고립되지 않고 핵심 개념·일반화에 연결되어 조직되는가\n' +
    '- retrieval: 한 번 보여주고 끝이 아니라, 떠올리게 하고(인출) 간격을 두는 장치가 있는가\n' +
    '- load: 학년 수준에 맞는 비계가 있고 신규 요소를 한꺼번에 쏟지 않는가\n' +
    '- metacog: 학습 목표 설정과 성찰의 기회가 명시되어 있는가\n' +
    '- assess: 형성평가가 과정 곳곳에 있고, 최종 평가가 일반화와 정렬되는가\n' +
    '- transfer: 수업에서 직접 다루지 않은 새로운 맥락에 적용하는 과제가 있는가\n' +
    '- phases: 관계 맺기~성찰하기 일곱 단계가 단원 흐름에 모두 있는가\n' +
    '- genquality: 일반화가 현재 시제이고, 시간·장소·인물에 매이지 않으며, 두 개 이상의 개념 관계를 담는가\n' +
    '- realism: ' + gradeStageText_(ctx.gradeBand) + ' 교실에서 단원 기간 안에 실제로 소화 가능한 분량인가\n' +
    'evidence에는 설계 속 구체적 요소의 이름(활동·전략·질문·과제명)을 인용해 1~2문장으로 근거를 써. advice는 warn일 때만 보완 제안 1문장, pass면 빈 문자열 "".\n\n' +
    '작업 ② 단원 개요: 단원을 시작할 때 교사가 학생들에게 직접 들려줄 소개문 후보를 3개 써줘. 조건:\n' +
    '- ' + gradeStageText_(ctx.gradeBand) + ' 눈높이의 친근한 존댓말, 3~5문장.\n' +
    '- 호기심을 끄는 참여 질문 1~2개로 시작하고, "이 단원에서 우리는 ……을 배울 것입니다/거예요" 형태의 문장을 포함해.\n' +
    '- 세 후보는 분위기가 서로 다르게(이야기형 / 질문형 / 도전형 등). note에는 교사용 한 줄 설명.\n\n' +
    '설명이나 마크다운 없이 아래 JSON 형식으로만 답해(check는 위 10개 id를 모두 포함):\n' +
    '{"check":[{"id":"precon","status":"pass","evidence":"","advice":""}],"overviews":[{"text":"","note":""}]}';

  return callGemini(prompt);
}

/* ══════════ 4. 내보내기 (다음 빌드: docx.js) ══════════ */
function authPing(){ return Promise.resolve({ ok: true }); }
function exportDesign(p){
  return Promise.reject(new Error('v3.0에서는 내보내기가 워드(.docx) 파일 다운로드 방식으로 바뀌며, 다음 업데이트에서 제공됩니다. 설계 내용은 브라우저 보관함에 안전하게 저장되어 있습니다.'));
}
function exportPdf(docId){
  return Promise.reject(new Error('PDF 내보내기는 다음 업데이트에서 제공됩니다.'));
}

/* ══════════ 5. google.script.run 호환 심(shim) ══════════ */
var __API = {
  getInitData: getInitData, getStandards: getStandards, getAreas: getAreas,
  getStrategyBank: function(){ return preloadBank_(); }, getTonghapData: getTonghapData,
  hasUserApiKey: function(){ return Promise.resolve(hasUserApiKey()); },
  saveUserApiKey: saveUserApiKey, clearUserApiKey: clearUserApiKey, whoAmI: whoAmI,
  generateLenses: generateLenses, generateTitles: generateTitles, generateCoreIdeas: generateCoreIdeas,
  generateStrands: generateStrands, generateStrandConcepts: generateStrandConcepts,
  generatePreconceptions: generatePreconceptions, generateGeneralizations: generateGeneralizations,
  generateQuestions: generateQuestions, generateGrasps: generateGrasps, generateRafts: generateRafts,
  generateFlowStrand: generateFlowStrand, generateStrategies: generateStrategies, generateReview: generateReview,
  authPing: authPing, exportDesign: exportDesign, exportPdf: exportPdf
};

function normErr_(e){
  if (e instanceof Error) return e;
  var err = new Error(msgOf_(e) || '알 수 없는 오류가 발생했습니다.');
  return err;
}

function Runner_(){ this._ok = null; this._fail = null; }
Runner_.prototype.withSuccessHandler = function(f){ this._ok = f; return this; };
Runner_.prototype.withFailureHandler = function(f){ this._fail = f; return this; };

function makeMethod_(name){
  return function(){
    var self = this, args = arguments;
    Promise.resolve().then(function(){
      return __API[name].apply(null, args);
    }).then(function(res){
      if (self._ok) { try { self._ok(res); } catch (e) { console.error('[shim:' + name + '] 성공 핸들러 오류', e); } }
    }, function(err){
      err = normErr_(err);
      if (self._fail) { try { self._fail(err); } catch (e2) { console.error('[shim:' + name + '] 실패 핸들러 오류', e2); } }
      else console.error('[shim:' + name + ']', err);
    });
    return undefined;
  };
}

var runBase = {};
(function(){
  for (var name in __API) {
    if (!__API.hasOwnProperty(name)) continue;
    Runner_.prototype[name] = makeMethod_(name);
    runBase[name] = (function(n){ return function(){ var r = new Runner_(); return r[n].apply(r, arguments); }; })(name);
  }
  runBase.withSuccessHandler = function(f){ return new Runner_().withSuccessHandler(f); };
  runBase.withFailureHandler = function(f){ return new Runner_().withFailureHandler(f); };
})();

window.google = window.google || {};
window.google.script = window.google.script || {};
window.google.script.run = runBase;

})();
