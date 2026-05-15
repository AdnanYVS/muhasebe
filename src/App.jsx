import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  TrendingUp, TrendingDown, Wallet, Receipt, Moon, Sun,
  Send, Filter, Trash2, PlusCircle, BarChart2, X, ClipboardList,
  Upload, Download, Landmark, ArrowUpCircle, ArrowDownCircle,
  CheckCircle2, RefreshCw, LayoutDashboard, Settings, Calendar,
  FileSpreadsheet, UserPlus, AlertTriangle, Target, FileText
} from 'lucide-react';
import {
  PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, LineChart, Line
} from 'recharts';

// ── Sabitler ────────────────────────────────────────────────
const DEFAULT_PERSONELLER = ['Personel 1', 'Personel 2', 'Personel 3'];
const KATEGORILER = ['Ofis', 'Üretim/Tarım', 'Pazarlama', 'Seyahat', 'Diğer'];
const KDV_ORANLARI = [0, 1, 10, 20];
const RENK_PALETI = ['#6366f1', '#22d3ee', '#f59e0b', '#10b981', '#f43f5e'];
const LS = { tx: 'muh_tx', kasa: 'muh_kasa', personeller: 'muh_personeller', butce: 'muh_butce', proforma: 'muh_proforma', musteriler: 'muh_musteriler', logo: 'muh_logo', sabitGiderler: 'muh_sabit' };
const BUTCE_DEFAULT = Object.fromEntries(KATEGORILER.map(k => [k, 0]));
const BIRIMLER = ['Adet', 'Kg', 'Ton', 'm²', 'Saat', 'Gün', 'Ay'];

const MOCK_DATA = [
  { id: 101, tarih: '2026-03-05', tur: 'Gelir', personel: null, aciklama: 'Toprak mahsulleri satışı', kategori: 'Üretim/Tarım', netTutar: 60000, kdvOrani: 1, kdvTutar: 600, toplamTutar: 60600 },
  { id: 102, tarih: '2026-03-10', tur: 'Gider', personel: 'Personel 1', aciklama: 'Ofis mobilyası', kategori: 'Ofis', netTutar: 8000, kdvOrani: 20, kdvTutar: 1600, toplamTutar: 9600 },
  { id: 103, tarih: '2026-03-15', tur: 'Gider', personel: 'Personel 2', aciklama: 'Reklam ajansı ödemesi', kategori: 'Pazarlama', netTutar: 7000, kdvOrani: 20, kdvTutar: 1400, toplamTutar: 8400 },
  { id: 201, tarih: '2026-04-03', tur: 'Gelir', personel: null, aciklama: 'Yazılım danışmanlığı', kategori: 'Diğer', netTutar: 45000, kdvOrani: 20, kdvTutar: 9000, toplamTutar: 54000 },
  { id: 202, tarih: '2026-04-08', tur: 'Gelir', personel: null, aciklama: 'Fide ihracatı', kategori: 'Üretim/Tarım', netTutar: 38000, kdvOrani: 1, kdvTutar: 380, toplamTutar: 38380 },
  { id: 203, tarih: '2026-04-12', tur: 'Gider', personel: 'Personel 3', aciklama: 'İstanbul iş seyahati', kategori: 'Seyahat', netTutar: 4500, kdvOrani: 10, kdvTutar: 450, toplamTutar: 4950 },
  { id: 204, tarih: '2026-04-18', tur: 'Gider', personel: 'Personel 1', aciklama: 'Sulama sistemi bakımı', kategori: 'Üretim/Tarım', netTutar: 12000, kdvOrani: 1, kdvTutar: 120, toplamTutar: 12120 },
  { id: 1, tarih: '2026-05-01', tur: 'Gelir', personel: null, aciklama: 'Müşteri ödemesi - Tahıl satışı', kategori: 'Üretim/Tarım', netTutar: 85000, kdvOrani: 1, kdvTutar: 850, toplamTutar: 85850 },
  { id: 2, tarih: '2026-05-02', tur: 'Gider', personel: 'Personel 1', aciklama: 'Ofis kırtasiye alımı', kategori: 'Ofis', netTutar: 1500, kdvOrani: 20, kdvTutar: 300, toplamTutar: 1800 },
  { id: 3, tarih: '2026-05-03', tur: 'Gider', personel: 'Personel 2', aciklama: 'Sosyal medya reklamı', kategori: 'Pazarlama', netTutar: 5000, kdvOrani: 20, kdvTutar: 1000, toplamTutar: 6000 },
  { id: 4, tarih: '2026-05-05', tur: 'Gelir', personel: null, aciklama: 'Danışmanlık hizmeti', kategori: 'Diğer', netTutar: 12000, kdvOrani: 20, kdvTutar: 2400, toplamTutar: 14400 },
  { id: 5, tarih: '2026-05-06', tur: 'Gider', personel: 'Personel 3', aciklama: 'Ankara iş seyahati', kategori: 'Seyahat', netTutar: 3200, kdvOrani: 10, kdvTutar: 320, toplamTutar: 3520 },
  { id: 6, tarih: '2026-05-08', tur: 'Gider', personel: 'Personel 1', aciklama: 'Gübre ve tarım ilacı', kategori: 'Üretim/Tarım', netTutar: 8500, kdvOrani: 1, kdvTutar: 85, toplamTutar: 8585 },
  { id: 7, tarih: '2026-05-10', tur: 'Gelir', personel: null, aciklama: 'Fide satışı - B2B', kategori: 'Üretim/Tarım', netTutar: 22000, kdvOrani: 1, kdvTutar: 220, toplamTutar: 22220 },
  { id: 8, tarih: '2026-05-12', tur: 'Gider', personel: 'Personel 2', aciklama: 'Yazıcı toner', kategori: 'Ofis', netTutar: 800, kdvOrani: 20, kdvTutar: 160, toplamTutar: 960 },
];

const MOCK_KASA = [
  { id: 9001, tarih: '2026-03-01', tur: 'Giriş', aciklama: 'İşletme açılış bakiyesi', tutar: 50000 },
];

// ── localStorage Yardımcıları ────────────────────────────────
function lsGet(key, fallback) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; }
  catch { return fallback; }
}
function lsSet(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
}

// ── Yardımcı Fonksiyonlar ────────────────────────────────────
function formatTL(n) {
  return new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY', maximumFractionDigits: 0 }).format(n);
}
function hesaplaKdv(net, oran) { return Math.round(net * (oran / 100)); }
function ayEtiketi(ym) {
  return new Intl.DateTimeFormat('tr-TR', { month: 'short', year: '2-digit' }).format(new Date(ym + '-01'));
}

// ── Tarih Aralığı Hesaplama ──────────────────────────────────
function calcDateRange(preset, from, to) {
  if (preset === 'tumu') return { from: '', to: '' };
  if (preset === 'ozel')  return { from, to };
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth();
  const today = now.toISOString().split('T')[0];
  const pad = n => String(n).padStart(2, '0');
  if (preset === 'buAy')     return { from: `${y}-${pad(m + 1)}-01`, to: today };
  if (preset === 'gecenAy')  return { from: new Date(y, m - 1, 1).toISOString().split('T')[0], to: new Date(y, m, 0).toISOString().split('T')[0] };
  if (preset === 'buCeyrek') return { from: new Date(y, Math.floor(m / 3) * 3, 1).toISOString().split('T')[0], to: today };
  if (preset === 'buYil')    return { from: `${y}-01-01`, to: today };
  return { from: '', to: '' };
}

// ── Excel İşlemleri ──────────────────────────────────────────
function parseExcelDate(val) {
  if (!val) return new Date().toISOString().split('T')[0];
  if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400000)).toISOString().split('T')[0];
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s); return isNaN(d) ? new Date().toISOString().split('T')[0] : d.toISOString().split('T')[0];
}

function findCol(headers, keys) {
  for (const k of keys) { const i = headers.findIndex(h => h.includes(k)); if (i !== -1) return i; }
  return -1;
}

function parseExcelRows(workbook, personellerList) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return { parsed: [], errors: [{ row: 1, reason: 'Dosya boş' }] };

  const hdr = rows[0].map(h => String(h).toLowerCase().trim());
  const col = {
    tarih:    findCol(hdr, ['tarih', 'date']),
    tur:      findCol(hdr, ['tür', 'tur', 'type']),
    personel: findCol(hdr, ['personel', 'personnel']),
    aciklama: findCol(hdr, ['açıklama', 'aciklama', 'description']),
    kategori: findCol(hdr, ['kategori', 'category']),
    netTutar: findCol(hdr, ['net tutar', 'nettutar', 'net', 'tutar', 'amount']),
    kdvOrani: findCol(hdr, ['kdv oranı', 'kdvorani', 'kdv %', 'kdv', 'vat']),
  };
  const parsed = [], errors = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(c => !c)) continue;
    try {
      const net = parseFloat(String(col.netTutar !== -1 ? row[col.netTutar] : row[5]).replace(/[^\d.-]/g, ''));
      if (!net || isNaN(net) || net <= 0) { errors.push({ row: i + 1, reason: 'Tutar geçersiz' }); continue; }
      const turRaw = String(col.tur !== -1 ? row[col.tur] : row[1]).toLowerCase();
      const tur = /gelir|income/.test(turRaw) ? 'Gelir' : 'Gider';
      const kdvRawStr = String(col.kdvOrani !== -1 ? row[col.kdvOrani] : row[6]).replace(/[^\d]/g, '');
      const kdvRaw = kdvRawStr !== '' ? parseInt(kdvRawStr) : null;
      const kdvOrani = kdvRaw !== null && KDV_ORANLARI.includes(kdvRaw) ? kdvRaw : 20;
      const personelRaw = String(col.personel !== -1 ? row[col.personel] : row[2]).trim();
      const personel = tur === 'Gider' ? (personellerList.find(p => personelRaw.toLowerCase().includes(p.toLowerCase())) || personellerList[0]) : null;
      const katRaw = String(col.kategori !== -1 ? row[col.kategori] : row[4]).trim();
      const kategori = KATEGORILER.find(k => k.toLowerCase() === katRaw.toLowerCase()) || 'Diğer';
      const acRaw = String(col.aciklama !== -1 ? row[col.aciklama] : row[3]).trim() || (tur === 'Gelir' ? 'Gelir kaydı' : 'Gider kaydı');
      const kdvTutar = hesaplaKdv(net, kdvOrani);
      parsed.push({ id: Date.now() + i, tarih: parseExcelDate(col.tarih !== -1 ? row[col.tarih] : row[0]), tur, personel, aciklama: acRaw.length > 60 ? acRaw.slice(0, 57) + '...' : acRaw, kategori, netTutar: net, kdvOrani, kdvTutar, toplamTutar: net + kdvTutar });
    } catch (e) { errors.push({ row: i + 1, reason: e.message }); }
  }
  return { parsed, errors };
}

function downloadTemplate() {
  const h = ['Tarih', 'Tür', 'Personel', 'Açıklama', 'Kategori', 'Net Tutar', 'KDV Oranı'];
  const d = [
    ['2026-01-05', 'Gelir', '', 'Tahıl satışı', 'Üretim/Tarım', 50000, 1],
    ['2026-01-10', 'Gider', 'Personel 1', 'Ofis kırtasiye', 'Ofis', 2000, 20],
    ['2026-02-03', 'Gelir', '', 'Danışmanlık', 'Diğer', 15000, 20],
    ['2026-02-15', 'Gider', 'Personel 2', 'Reklam ajansı', 'Pazarlama', 3500, 20],
  ];
  const ws = XLSX.utils.aoa_to_sheet([h, ...d]);
  ws['!cols'] = [12, 8, 14, 32, 16, 12, 10].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Muhasebe');
  XLSX.writeFile(wb, 'muhasebe_sablon.xlsx');
}

function exportMuhasebeExcel(rows, etiket = '') {
  const h = ['Tarih', 'Tür', 'Personel', 'Açıklama', 'Kategori', 'Net Tutar (₺)', 'KDV %', 'KDV Tutarı (₺)', 'Toplam (₺)'];
  const data = rows.map(r => [r.tarih, r.tur, r.personel || '', r.aciklama, r.kategori, r.netTutar, r.kdvOrani, r.kdvTutar, r.toplamTutar]);
  const gelir = rows.filter(r => r.tur === 'Gelir');
  const gider = rows.filter(r => r.tur === 'Gider');
  data.push([], ['', '', '', '', 'ÖZET', '', '', '', ''],
    ['', '', '', '', 'Toplam Gelir (net)', gelir.reduce((s, r) => s + r.netTutar, 0)],
    ['', '', '', '', 'Toplam Gider (net)', gider.reduce((s, r) => s + r.netTutar, 0)],
    ['', '', '', '', 'Net Kâr / Zarar', gelir.reduce((s, r) => s + r.netTutar, 0) - gider.reduce((s, r) => s + r.netTutar, 0)],
    ['', '', '', '', 'Alınan KDV', gelir.reduce((s, r) => s + r.kdvTutar, 0)],
    ['', '', '', '', 'Ödenen KDV', gider.reduce((s, r) => s + r.kdvTutar, 0)],
    ['', '', '', '', 'Ödenecek KDV', gelir.reduce((s, r) => s + r.kdvTutar, 0) - gider.reduce((s, r) => s + r.kdvTutar, 0)],
  );
  const ws = XLSX.utils.aoa_to_sheet([h, ...data]);
  ws['!cols'] = [12, 8, 14, 30, 14, 14, 8, 14, 14].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Muhasebe');
  XLSX.writeFile(wb, `muhasebe_rapor${etiket ? '_' + etiket : ''}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function exportKasaExcel(hareketler) {
  let bal = 0;
  const sirali = [...hareketler].sort((a, b) => new Date(a.tarih) - new Date(b.tarih)).map(h => {
    if (h.tur === 'Giriş' || h.tur === 'Kar Devri') bal += h.tutar; else bal -= h.tutar;
    return [h.tarih, h.aciklama, h.tur, (h.tur === 'Giriş' || h.tur === 'Kar Devri') ? h.tutar : -h.tutar, bal];
  });
  const ws = XLSX.utils.aoa_to_sheet([['Tarih', 'Açıklama', 'Tür', 'Tutar (₺)', 'Bakiye (₺)'], ...sirali]);
  ws['!cols'] = [12, 35, 12, 14, 14].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Kasa');
  XLSX.writeFile(wb, `kasa_ozeti_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ── Doğal Dil Parser ─────────────────────────────────────────
async function askGemini(userMessage, transactions, personellerList) {
  const key = import.meta.env.VITE_GROQ_API_KEY;
  const toplamGelir = transactions.filter(t => t.tur === 'Gelir').reduce((s, t) => s + t.netTutar, 0);
  const toplamGider = transactions.filter(t => t.tur === 'Gider').reduce((s, t) => s + t.netTutar, 0);

  const systemPrompt = `Sen bir muhasebe asistanısın. Türkçe konuşursun.

Mevcut durum:
- Toplam kayıt: ${transactions.length}
- Toplam gelir: ${toplamGelir} TL
- Toplam gider: ${toplamGider} TL
- Net kâr: ${toplamGelir - toplamGider} TL
- Personel: ${personellerList.join(', ')}
- Kategoriler: Ofis, Üretim/Tarım, Pazarlama, Seyahat, Diğer

Türkiye KDV oranı kuralları:
- Kira, aidat, kira ödemesi: %0 (KDV yok)
- Temel gıda, tarım ürünleri, fide, tohum, gübre: %1
- Tekstil, giyim, ulaşım, seyahat: %10
- Genel hizmet, danışmanlık, elektronik, ofis malzemesi, reklam: %20
- Kullanıcı belirtmediyse kira/aidat türü işlemlerde %0 kullan

Kullanıcı yeni bir işlem girmek istiyorsa YALNIZCA şu JSON'u döndür (başka hiç metin ekleme):
{"tip":"islem","tur":"Gelir","aciklama":"...","netTutar":1000,"kdvOrani":20,"kategori":"Ofis","personel":"Personel 1"}

kdvOrani değeri: 0, 1, 10 veya 20 olabilir. Kira için mutlaka 0 kullan.
Soru soruyorsa kısa ve net Türkçe yanıt ver.`;

  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userMessage }],
      temperature: 0.1, max_tokens: 512
    })
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.code === 'rate_limit_exceeded' ? 'QUOTA' : data.error.message);
  const text = (data.choices?.[0]?.message?.content || '').trim();

  try {
    const clean = text.replace(/```json|```/g, '').trim();
    const json = JSON.parse(clean);
    if (json.tip === 'islem') {
      const net = Number(json.netTutar) || 0;
      const kdvOrani = (json.kdvOrani !== undefined && json.kdvOrani !== null) ? Number(json.kdvOrani) : 20;
      const kdvTutar = hesaplaKdv(net, kdvOrani);
      return { type: 'transaction', data: { id: Date.now(), tarih: new Date().toISOString().split('T')[0], tur: json.tur || 'Gider', personel: json.personel || (json.tur !== 'Gelir' ? personellerList[0] : null), aciklama: json.aciklama || userMessage, kategori: json.kategori || 'Diğer', netTutar: net, kdvOrani, kdvTutar, toplamTutar: net + kdvTutar } };
    }
  } catch (_) {}
  return { type: 'answer', text };
}

function parseChat(metin, personellerList) {
  const text = metin.toLowerCase();
  const tur = /gelir|satış|tahsil|kazanç|ödeme aldık|müşteriden/.test(text) ? 'Gelir' : 'Gider';
  const tutarMatch = metin.match(/(\d[\d.,]*)\s*(?:tl|lira|₺)/i);
  if (!tutarMatch) return null;
  const rawTutar = parseFloat(tutarMatch[1].replace(/\./g, '').replace(',', '.'));
  let kdvOrani = 20;
  if (/kira|aidat|rent/.test(text)) kdvOrani = 0;
  else if (/%\s*0\b|kdv\s*yok|kdvsiz/.test(text)) kdvOrani = 0;
  else if (/%\s*1\b|%1|kdv\s*1/.test(text)) kdvOrani = 1;
  else if (/%\s*10\b|%10|kdv\s*10/.test(text)) kdvOrani = 10;
  const kdvDahil = /kdv\s*dahil|kdv'li|toplam/.test(text);
  let netTutar, kdvTutar;
  if (kdvDahil) { netTutar = Math.round(rawTutar / (1 + kdvOrani / 100)); kdvTutar = rawTutar - netTutar; }
  else { netTutar = rawTutar; kdvTutar = hesaplaKdv(netTutar, kdvOrani); }
  let personel = null;
  if (tur === 'Gider') {
    for (const p of personellerList) { if (text.includes(p.toLowerCase())) { personel = p; break; } }
    if (!personel) personel = personellerList[0] || null;
  }
  let kategori = 'Diğer';
  if (/ofis|kırtasiye|toner|yazıcı/.test(text)) kategori = 'Ofis';
  else if (/tarım|üretim|gübre|fide|tohum/.test(text)) kategori = 'Üretim/Tarım';
  else if (/reklam|pazarlama|sosyal medya/.test(text)) kategori = 'Pazarlama';
  else if (/seyahat|uçak|otel/.test(text)) kategori = 'Seyahat';
  return { id: Date.now(), tarih: new Date().toISOString().split('T')[0], tur, personel, aciklama: metin.length > 60 ? metin.slice(0, 57) + '...' : metin, kategori, netTutar, kdvOrani, kdvTutar, toplamTutar: netTutar + kdvTutar };
}

// ── Proforma Yardımcıları ────────────────────────────────────
function nextFaturaNo(proformalar) {
  const y = new Date().getFullYear();
  const nums = proformalar.map(p => { const m = p.faturaNo.match(/PRF-\d{4}-(\d+)/); return m ? parseInt(m[1]) : 0; });
  return `PRF-${y}-${String((nums.length ? Math.max(...nums) : 0) + 1).padStart(3, '0')}`;
}
function calcKalem(k) {
  const net = k.miktar * k.birimFiyat;
  const kdv = hesaplaKdv(net, k.kdvOrani);
  return { ...k, netTutar: net, kdvTutar: kdv, toplam: net + kdv };
}
function exportProformaExcel(p) {
  const kalemler = p.kalemler.map(calcKalem);
  const toplamNet = kalemler.reduce((s, k) => s + k.netTutar, 0);
  const toplamKdv = kalemler.reduce((s, k) => s + k.kdvTutar, 0);
  const h = ['#', 'Açıklama', 'Miktar', 'Birim', 'Birim Fiyat (₺)', 'KDV %', 'KDV Tutarı (₺)', 'Toplam (₺)'];
  const rows = kalemler.map((k, i) => [i + 1, k.aciklama, k.miktar, k.birim, k.birimFiyat, k.kdvOrani, k.kdvTutar, k.toplam]);
  rows.push([], ['', '', '', '', 'Ara Toplam', '', '', toplamNet], ['', '', '', '', 'KDV', '', '', toplamKdv], ['', '', '', '', 'GENEL TOPLAM', '', '', toplamNet + toplamKdv]);
  const ws = XLSX.utils.aoa_to_sheet([[`PROFORMA FATURA — ${p.faturaNo}`], [`Müşteri: ${p.musteri.ad}${p.musteri.vergiNo ? ' | VKN: ' + p.musteri.vergiNo : ''}`], [`Tarih: ${p.tarih} | Geçerlilik: ${p.gecerlilikTarihi}`], [], h, ...rows]);
  ws['!cols'] = [4, 30, 8, 8, 14, 8, 14, 14].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Proforma');
  XLSX.writeFile(wb, `${p.faturaNo}_${p.musteri.ad.replace(/\s+/g, '_')}.xlsx`);
}

// ── Temel Bileşenler ─────────────────────────────────────────
function Card({ children, className = '' }) {
  return <div className={`rounded-2xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-sm p-5 ${className}`}>{children}</div>;
}
function Badge({ children, color = 'blue' }) {
  const colors = { blue: 'bg-blue-100 text-blue-700 dark:bg-blue-900 dark:text-blue-300', green: 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900 dark:text-emerald-300', red: 'bg-rose-100 text-rose-700 dark:bg-rose-900 dark:text-rose-300', amber: 'bg-amber-100 text-amber-700 dark:bg-amber-900 dark:text-amber-300', violet: 'bg-violet-100 text-violet-700 dark:bg-violet-900 dark:text-violet-300' };
  return <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${colors[color]}`}>{children}</span>;
}

// ── Tarih Filtresi Çubuğu ────────────────────────────────────
function DateFilterBar({ preset, from, to, onPreset, onFrom, onTo, count, total }) {
  const PRESETS = [
    { id: 'tumu', label: 'Tümü' },
    { id: 'buAy', label: 'Bu Ay' },
    { id: 'gecenAy', label: 'Geçen Ay' },
    { id: 'buCeyrek', label: 'Bu Çeyrek' },
    { id: 'buYil', label: 'Bu Yıl' },
    { id: 'ozel', label: 'Özel' },
  ];
  return (
    <div className="flex flex-wrap items-center gap-2">
      <Calendar size={13} className="text-slate-400 shrink-0" />
      {PRESETS.map(p => (
        <button key={p.id} onClick={() => onPreset(p.id)}
          className={`text-xs px-3 py-1.5 rounded-lg border font-medium transition-colors whitespace-nowrap ${preset === p.id ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300'}`}>
          {p.label}
        </button>
      ))}
      {preset === 'ozel' && (<>
        <input type="date" value={from} onChange={e => onFrom(e.target.value)}
          className="text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-1.5 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
        <span className="text-xs text-slate-400">—</span>
        <input type="date" value={to} onChange={e => onTo(e.target.value)}
          className="text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-1.5 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
      </>)}
      {preset !== 'tumu' && (
        <span className="text-xs text-indigo-600 dark:text-indigo-400 font-medium">
          {count} / {total} kayıt
        </span>
      )}
    </div>
  );
}

// ── Manuel Form Modal ────────────────────────────────────────
function FormModal({ onClose, onAdd, personeller }) {
  const [form, setForm] = useState({ tarih: new Date().toISOString().split('T')[0], tur: 'Gider', personel: personeller[0] || '', aciklama: '', kategori: KATEGORILER[0], netTutar: '', kdvOrani: 20, kdvDahil: false, faturaNo: '' });
  const raw = parseFloat(form.netTutar) || 0;
  const kdvTutar = form.kdvDahil ? raw - Math.round(raw / (1 + form.kdvOrani / 100)) : hesaplaKdv(raw, form.kdvOrani);
  const netGercek = form.kdvDahil ? raw - kdvTutar : raw;
  const toplam = netGercek + kdvTutar;
  const ic = 'w-full text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';
  function handleSubmit(e) {
    e.preventDefault();
    if (!raw || raw <= 0) return;
    onAdd({ id: Date.now(), tarih: form.tarih, tur: form.tur, personel: form.tur === 'Gelir' ? null : form.personel, aciklama: form.aciklama.trim() || (form.tur === 'Gelir' ? 'Gelir kaydı' : 'Gider kaydı'), kategori: form.kategori, netTutar: netGercek, kdvOrani: form.kdvOrani, kdvTutar, toplamTutar: toplam, ...(form.faturaNo.trim() && { faturaNo: form.faturaNo.trim() }) });
    onClose();
  }
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
      <div className="w-full max-w-md rounded-2xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-2xl">
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 dark:border-slate-700">
          <h2 className="text-sm font-semibold flex items-center gap-2"><ClipboardList size={15} className="text-indigo-500" /> Manuel Veri Girişi</h2>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700"><X size={14} /></button>
        </div>
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div className="grid grid-cols-2 gap-3">
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Tarih</label><input type="date" value={form.tarih} onChange={e => setForm(f => ({ ...f, tarih: e.target.value }))} className={ic} /></div>
            <div>
              <label className="text-xs font-medium text-slate-500 mb-1.5 block">Tür</label>
              <div className="flex gap-2">
                {['Gelir', 'Gider'].map(t => (<button key={t} type="button" onClick={() => setForm(f => ({ ...f, tur: t }))} className={`flex-1 text-xs py-2 rounded-lg border font-medium transition-colors ${form.tur === t ? (t === 'Gelir' ? 'bg-emerald-500 text-white border-emerald-500' : 'bg-rose-500 text-white border-rose-500') : 'border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700'}`}>{t}</button>))}
              </div>
            </div>
          </div>
          <div className="grid grid-cols-3 gap-3">
            <div className="col-span-2"><label className="text-xs font-medium text-slate-500 mb-1.5 block">Açıklama</label><input type="text" value={form.aciklama} onChange={e => setForm(f => ({ ...f, aciklama: e.target.value }))} placeholder="İşlem açıklaması..." className={ic} /></div>
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Fatura No <span className="font-normal text-slate-400">(opsiyonel)</span></label><input type="text" value={form.faturaNo} onChange={e => setForm(f => ({ ...f, faturaNo: e.target.value }))} placeholder="FT-001" className={ic} /></div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Kategori</label><select value={form.kategori} onChange={e => setForm(f => ({ ...f, kategori: e.target.value }))} className={ic}>{KATEGORILER.map(k => <option key={k}>{k}</option>)}</select></div>
            {form.tur === 'Gider' && (<div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Personel</label><select value={form.personel} onChange={e => setForm(f => ({ ...f, personel: e.target.value }))} className={ic}>{personeller.map(p => <option key={p}>{p}</option>)}</select></div>)}
          </div>
          <div className="grid grid-cols-3 gap-3">
            <div className="col-span-2"><label className="text-xs font-medium text-slate-500 mb-1.5 block">Tutar (₺) <span className="font-normal text-slate-400">{form.kdvDahil ? '— KDV dahil' : '— KDV hariç'}</span></label><input type="number" min="0" step="0.01" value={form.netTutar} onChange={e => setForm(f => ({ ...f, netTutar: e.target.value }))} placeholder="0" className={ic} /></div>
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">KDV</label><select value={form.kdvOrani} onChange={e => setForm(f => ({ ...f, kdvOrani: parseInt(e.target.value) }))} className={ic}>{KDV_ORANLARI.map(o => <option key={o} value={o}>%{o}</option>)}</select></div>
          </div>
          <label className="flex items-center gap-2 cursor-pointer select-none"><input type="checkbox" checked={form.kdvDahil} onChange={e => setForm(f => ({ ...f, kdvDahil: e.target.checked }))} className="w-4 h-4 accent-indigo-600" /><span className="text-xs text-slate-500">Girilen tutar KDV dahil</span></label>
          {raw > 0 && (<div className="rounded-xl bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 p-3 grid grid-cols-3 gap-2 text-center">
            <div><p className="text-xs text-slate-400 mb-0.5">Net</p><p className="text-sm font-semibold">{formatTL(netGercek)}</p></div>
            <div><p className="text-xs text-slate-400 mb-0.5">KDV %{form.kdvOrani}</p><p className="text-sm font-semibold text-amber-600 dark:text-amber-400">{formatTL(kdvTutar)}</p></div>
            <div><p className="text-xs text-slate-400 mb-0.5">Toplam</p><p className="text-sm font-bold text-indigo-600 dark:text-indigo-400">{formatTL(toplam)}</p></div>
          </div>)}
          <div className="flex gap-2 pt-1">
            <button type="button" onClick={onClose} className="flex-1 text-xs py-2.5 rounded-xl border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">İptal</button>
            <button type="submit" disabled={!raw || raw <= 0} className="flex-1 text-xs py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white font-medium">Kaydet</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ── Excel Önizleme Modal ─────────────────────────────────────
function ExcelPreviewModal({ parsed, errors, onConfirm, onClose }) {
  const [sel, setSel] = useState(() => new Set(parsed.map((_, i) => i)));
  const toggle = i => setSel(s => { const n = new Set(s); n.has(i) ? n.delete(i) : n.add(i); return n; });
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
      <div className="w-full max-w-2xl max-h-[90vh] flex flex-col rounded-2xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-2xl">
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 dark:border-slate-700">
          <div><h2 className="text-sm font-semibold flex items-center gap-2"><Upload size={14} className="text-indigo-500" /> Excel Önizlemesi</h2><p className="text-xs text-slate-400 mt-0.5">{parsed.length} kayıt okundu{errors.length > 0 && `, ${errors.length} satır hatalı (atlandı)`}</p></div>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700"><X size={14} /></button>
        </div>
        {errors.length > 0 && <div className="mx-5 mt-3 p-3 rounded-xl bg-rose-50 dark:bg-rose-900/20 border border-rose-200 dark:border-rose-800 text-xs text-rose-700 dark:text-rose-300">Atlanan satırlar: {errors.map(e => `Satır ${e.row} (${e.reason})`).join(' • ')}</div>}
        {parsed.length === 0 ? (
          <div className="flex-1 flex items-center justify-center text-sm text-slate-400 p-8">İçe aktarılabilir kayıt bulunamadı. Şablonu indirip doldurun.</div>
        ) : (
          <div className="flex-1 overflow-auto mx-5 my-3 rounded-xl border border-slate-200 dark:border-slate-600">
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                <tr className="text-slate-500">
                  <th className="px-3 py-2"><input type="checkbox" checked={sel.size === parsed.length} onChange={() => setSel(sel.size === parsed.length ? new Set() : new Set(parsed.map((_, i) => i)))} className="accent-indigo-600" /></th>
                  {['Tarih', 'Tür', 'Personel', 'Açıklama', 'Net Tutar', 'KDV'].map(h => <th key={h} className="px-3 py-2 text-left font-medium">{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {parsed.map((r, i) => (
                  <tr key={i} onClick={() => toggle(i)} className={`border-t border-slate-100 dark:border-slate-700 cursor-pointer ${!sel.has(i) ? 'opacity-40' : ''}`}>
                    <td className="px-3 py-2"><input type="checkbox" checked={sel.has(i)} onChange={() => toggle(i)} onClick={e => e.stopPropagation()} className="accent-indigo-600" /></td>
                    <td className="px-3 py-2 text-slate-500 whitespace-nowrap">{r.tarih}</td>
                    <td className="px-3 py-2"><Badge color={r.tur === 'Gelir' ? 'green' : 'red'}>{r.tur}</Badge></td>
                    <td className="px-3 py-2 text-slate-500">{r.personel || '—'}</td>
                    <td className="px-3 py-2 max-w-[130px] truncate" title={r.aciklama}>{r.aciklama}</td>
                    <td className="px-3 py-2 font-medium">{formatTL(r.netTutar)}</td>
                    <td className="px-3 py-2"><Badge color="amber">%{r.kdvOrani}</Badge></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div className="flex gap-2 px-5 pb-4">
          <button onClick={onClose} className="flex-1 text-xs py-2.5 rounded-xl border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">İptal</button>
          <button onClick={() => onConfirm(parsed.filter((_, i) => sel.has(i)))} disabled={sel.size === 0} className="flex-1 text-xs py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white font-medium">{sel.size} Kaydı İçe Aktar</button>
        </div>
      </div>
    </div>
  );
}

// ── Sabit Giderler Kartı ─────────────────────────────────────
function SabitGiderlerCard({ sabitGiderler, onEkle, onSil, onUygula, transactions }) {
  const buAy = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
  const buAyEtiket = new Intl.DateTimeFormat('tr-TR', { month: 'long', year: 'numeric' }).format(new Date());
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ ad: '', tutar: '', kdvOrani: 0, kategori: 'Ofis' });
  const ic = 'w-full text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';

  const uygulanmislar = new Set(
    transactions.filter(t => t.sabitGiderId && t.tarih.startsWith(buAy)).map(t => t.sabitGiderId)
  );
  const bekleyenler = sabitGiderler.filter(sg => !uygulanmislar.has(sg.id));

  function handleEkle(e) {
    e.preventDefault();
    const tutar = parseFloat(form.tutar);
    if (!form.ad.trim() || !tutar || tutar <= 0) return;
    onEkle({ id: Date.now(), ad: form.ad.trim(), tutar, kdvOrani: form.kdvOrani, kategori: form.kategori });
    setForm({ ad: '', tutar: '', kdvOrani: 0, kategori: 'Ofis' });
    setShowForm(false);
  }

  return (
    <Card>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-sm font-semibold flex items-center gap-2">
          <Calendar size={15} className="text-violet-500" /> Sabit Giderler
          <span className="text-xs font-normal text-slate-400">(aylık tekrarlayan)</span>
        </h2>
        <button onClick={() => setShowForm(s => !s)} className="text-xs px-2.5 py-1.5 rounded-lg bg-violet-50 dark:bg-violet-900/30 text-violet-600 dark:text-violet-400 hover:bg-violet-100 dark:hover:bg-violet-900/50 font-medium flex items-center gap-1">
          <PlusCircle size={11} /> Tanımla
        </button>
      </div>

      {bekleyenler.length > 0 && (
        <div className="mb-3 p-3 rounded-xl bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-700 flex items-start gap-2">
          <AlertTriangle size={14} className="text-amber-500 shrink-0 mt-0.5" />
          <p className="text-xs text-amber-700 dark:text-amber-300 font-medium">
            {buAyEtiket} için <strong>{bekleyenler.length}</strong> sabit gider henüz tabloya eklenmedi.
          </p>
        </div>
      )}

      {showForm && (
        <form onSubmit={handleEkle} className="mb-4 p-3 rounded-xl border border-violet-200 dark:border-violet-700 bg-violet-50/50 dark:bg-violet-900/10 space-y-2">
          <p className="text-xs font-medium text-slate-600 dark:text-slate-300">Yeni Sabit Gider Tanımla</p>
          <div className="grid grid-cols-2 gap-2">
            <input value={form.ad} onChange={e => setForm(f => ({ ...f, ad: e.target.value }))} placeholder="Gider adı (ör: Ofis Kirası)" required className={ic} />
            <input type="number" min="0" step="0.01" value={form.tutar} onChange={e => setForm(f => ({ ...f, tutar: e.target.value }))} placeholder="Aylık tutar (₺)" className={ic} />
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <select value={form.kdvOrani} onChange={e => setForm(f => ({ ...f, kdvOrani: parseInt(e.target.value) }))} className={ic}>
                {KDV_ORANLARI.map(o => <option key={o} value={o}>KDV %{o}{o === 0 ? ' (muaf)' : ''}</option>)}
              </select>
            </div>
            <select value={form.kategori} onChange={e => setForm(f => ({ ...f, kategori: e.target.value }))} className={ic}>
              {KATEGORILER.map(k => <option key={k}>{k}</option>)}
            </select>
          </div>
          <div className="flex gap-2 justify-end">
            <button type="button" onClick={() => setShowForm(false)} className="text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">İptal</button>
            <button type="submit" className="text-xs px-3 py-1.5 rounded-lg bg-violet-600 hover:bg-violet-700 text-white font-medium">Kaydet</button>
          </div>
        </form>
      )}

      {sabitGiderler.length === 0 ? (
        <p className="text-xs text-slate-400 text-center py-4">Tanımlı sabit gider yok. Kira, abonelik gibi aylık tekrarlayan giderlerinizi ekleyin.</p>
      ) : (
        <div className="space-y-2">
          {sabitGiderler.map(sg => {
            const uygulandimi = uygulanmislar.has(sg.id);
            const kdvTutar = hesaplaKdv(sg.tutar, sg.kdvOrani);
            return (
              <div key={sg.id} className={`flex items-center justify-between p-3 rounded-xl border transition-colors ${uygulandimi ? 'border-slate-100 dark:border-slate-700 bg-slate-50/50 dark:bg-slate-800/20 opacity-60' : 'border-slate-200 dark:border-slate-600 hover:border-violet-300 dark:hover:border-violet-600'}`}>
                <div>
                  <span className="text-sm font-medium">{sg.ad}</span>
                  <div className="flex gap-3 mt-0.5">
                    <span className="text-xs text-slate-500">{formatTL(sg.tutar)}</span>
                    <Badge color={sg.kdvOrani === 0 ? 'green' : 'amber'}>KDV %{sg.kdvOrani}</Badge>
                    <Badge color="blue">{sg.kategori}</Badge>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  {uygulandimi ? (
                    <span className="flex items-center gap-1 text-xs text-emerald-600 dark:text-emerald-400 font-medium"><CheckCircle2 size={13} /> Eklendi</span>
                  ) : (
                    <button onClick={() => onUygula(sg, buAy)} className="text-xs px-2.5 py-1.5 rounded-lg bg-violet-600 hover:bg-violet-700 text-white font-medium whitespace-nowrap">
                      Bu Aya Ekle
                    </button>
                  )}
                  <button onClick={() => onSil(sg.id)} className="p-1 rounded text-slate-300 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"><X size={12} /></button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </Card>
  );
}

// ── Ofis Giderleri Kartı ─────────────────────────────────────
function OfisGiderleriCard({ onAdd }) {
  const today = new Date().toISOString().split('T')[0];
  const [form, setForm] = useState({ aciklama: '', tutar: '', kdvOrani: 20, kdvDahil: false, tarih: today, kategori: 'Ofis' });
  const ic = 'w-full text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';
  const raw = parseFloat(form.tutar) || 0;
  const kdvTutar = form.kdvDahil ? raw - Math.round(raw / (1 + form.kdvOrani / 100)) : hesaplaKdv(raw, form.kdvOrani);
  const netGercek = form.kdvDahil ? raw - kdvTutar : raw;

  function handleSubmit(e) {
    e.preventDefault();
    if (!raw || raw <= 0 || !form.aciklama.trim()) return;
    onAdd({
      id: Date.now(), tarih: form.tarih, tur: 'Gider', personel: null,
      aciklama: form.aciklama.trim(), kategori: form.kategori,
      netTutar: netGercek, kdvOrani: form.kdvOrani, kdvTutar, toplamTutar: netGercek + kdvTutar,
    });
    setForm({ aciklama: '', tutar: '', kdvOrani: 20, kdvDahil: false, tarih: today, kategori: 'Ofis' });
  }

  return (
    <Card>
      <h2 className="text-sm font-semibold mb-3 flex items-center gap-2">
        <Wallet size={15} className="text-teal-500" /> Ofis / Genel Gider
        <span className="text-xs font-normal text-slate-400">(personelsiz hızlı giriş)</span>
      </h2>
      <form onSubmit={handleSubmit} className="space-y-3">
        <div className="grid grid-cols-2 gap-2">
          <div className="col-span-2">
            <input value={form.aciklama} onChange={e => setForm(f => ({ ...f, aciklama: e.target.value }))} placeholder="Gider açıklaması..." required className={ic} />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <input type="number" min="0" step="0.01" value={form.tutar} onChange={e => setForm(f => ({ ...f, tutar: e.target.value }))} placeholder="Tutar (₺)" className={ic} />
          <select value={form.kdvOrani} onChange={e => setForm(f => ({ ...f, kdvOrani: parseInt(e.target.value) }))} className={ic}>
            {KDV_ORANLARI.map(o => <option key={o} value={o}>KDV %{o}</option>)}
          </select>
          <select value={form.kategori} onChange={e => setForm(f => ({ ...f, kategori: e.target.value }))} className={ic}>
            {KATEGORILER.map(k => <option key={k}>{k}</option>)}
          </select>
        </div>
        <div className="flex items-center gap-3">
          <input type="date" value={form.tarih} onChange={e => setForm(f => ({ ...f, tarih: e.target.value }))} className="text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500" />
          <label className="flex items-center gap-1.5 cursor-pointer select-none">
            <input type="checkbox" checked={form.kdvDahil} onChange={e => setForm(f => ({ ...f, kdvDahil: e.target.checked }))} className="w-3.5 h-3.5 accent-indigo-600" />
            <span className="text-xs text-slate-500">KDV dahil</span>
          </label>
          {raw > 0 && (
            <span className="text-xs text-slate-400 ml-auto">Net: <strong className="text-slate-600 dark:text-slate-200">{formatTL(netGercek)}</strong> + KDV: <strong className="text-amber-600">{formatTL(kdvTutar)}</strong></span>
          )}
        </div>
        <button type="submit" disabled={!raw || !form.aciklama.trim()} className="w-full text-xs py-2.5 rounded-xl bg-teal-600 hover:bg-teal-700 disabled:opacity-40 text-white font-medium flex items-center justify-center gap-1.5">
          <PlusCircle size={13} /> Gider Tablosuna Ekle
        </button>
      </form>
    </Card>
  );
}

// ── Kasa Sekmesi ─────────────────────────────────────────────
function KasaTab({ kasaHareketleri, aylikTrend, onKasaEkle, onKasaDelete }) {
  const [form, setForm] = useState({ tur: 'Giriş', aciklama: '', tutar: '', tarih: new Date().toISOString().split('T')[0] });
  const sirali = useMemo(() => {
    let bal = 0;
    return [...kasaHareketleri].sort((a, b) => new Date(a.tarih) - new Date(b.tarih) || a.id - b.id).map(h => { if (h.tur === 'Giriş' || h.tur === 'Kar Devri') bal += h.tutar; else bal -= h.tutar; return { ...h, bakiye: bal }; }).reverse();
  }, [kasaHareketleri]);
  const bakiye = sirali.length > 0 ? sirali[0].bakiye : 0;
  const devirAylar = useMemo(() => new Set(kasaHareketleri.filter(h => h.ayRef).map(h => h.ayRef)), [kasaHareketleri]);
  const ic = 'w-full text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';
  function handleSubmit(e) {
    e.preventDefault();
    const t = parseFloat(form.tutar);
    if (!t || t <= 0 || !form.aciklama.trim()) return;
    onKasaEkle({ id: Date.now(), tarih: form.tarih, tur: form.tur, aciklama: form.aciklama.trim(), tutar: t });
    setForm(f => ({ ...f, aciklama: '', tutar: '' }));
  }
  function handleDevir(ym, kar) {
    onKasaEkle({ id: Date.now(), tarih: new Date().toISOString().split('T')[0], tur: kar >= 0 ? 'Kar Devri' : 'Zarar Devri', aciklama: `${ayEtiketi(ym)} dönemi net ${kar >= 0 ? 'kâr' : 'zarar'} devri`, tutar: Math.abs(kar), ayRef: ym });
  }
  return (
    <div className="space-y-6">
      <Card>
        <div className="flex items-start justify-between">
          <div>
            <p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide flex items-center gap-1.5"><Landmark size={12} /> Kasa / Banka Bakiyesi</p>
            <p className={`text-4xl font-bold mt-2 ${bakiye >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600'}`}>{formatTL(bakiye)}</p>
            {sirali.length > 0 && <p className="text-xs text-slate-400 mt-1.5">Son: {sirali[0].tarih} — {sirali[0].aciklama}</p>}
          </div>
          <div className="flex gap-2">
            <button onClick={() => exportKasaExcel(kasaHareketleri)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors font-medium">
              <FileSpreadsheet size={13} /> Özet İndir
            </button>
            <div className={`w-12 h-12 rounded-2xl flex items-center justify-center ${bakiye >= 0 ? 'bg-emerald-100 dark:bg-emerald-900/40' : 'bg-rose-100 dark:bg-rose-900/40'}`}>
              <Landmark size={22} className={bakiye >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-500'} />
            </div>
          </div>
        </div>
      </Card>
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <Card>
          <h2 className="text-sm font-semibold mb-4 flex items-center gap-2">{form.tur === 'Giriş' ? <ArrowUpCircle size={15} className="text-emerald-500" /> : <ArrowDownCircle size={15} className="text-rose-500" />} Kasa Hareketi Ekle</h2>
          <form onSubmit={handleSubmit} className="space-y-3">
            <div className="flex gap-2">
              {['Giriş', 'Çıkış'].map(t => (<button key={t} type="button" onClick={() => setForm(f => ({ ...f, tur: t }))} className={`flex-1 text-xs py-2 rounded-lg border font-medium transition-colors flex items-center justify-center gap-1 ${form.tur === t ? (t === 'Giriş' ? 'bg-emerald-500 text-white border-emerald-500' : 'bg-rose-500 text-white border-rose-500') : 'border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700'}`}>{t === 'Giriş' ? <ArrowUpCircle size={11} /> : <ArrowDownCircle size={11} />} {t}</button>))}
            </div>
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Açıklama <span className="text-rose-400">*</span></label><input type="text" value={form.aciklama} onChange={e => setForm(f => ({ ...f, aciklama: e.target.value }))} placeholder={form.tur === 'Giriş' ? 'Para nereden geldi?' : 'Para nereye gitti?'} required className={ic} /></div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Tutar (₺)</label><input type="number" min="0" step="0.01" value={form.tutar} onChange={e => setForm(f => ({ ...f, tutar: e.target.value }))} placeholder="0" className={ic} /></div>
              <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Tarih</label><input type="date" value={form.tarih} onChange={e => setForm(f => ({ ...f, tarih: e.target.value }))} className={ic} /></div>
            </div>
            <button type="submit" disabled={!parseFloat(form.tutar) || parseFloat(form.tutar) <= 0 || !form.aciklama.trim()} className="w-full text-xs py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white font-medium">Hareketi Kaydet</button>
          </form>
        </Card>
        <Card className="lg:col-span-2">
          <h2 className="text-sm font-semibold mb-1 flex items-center gap-2"><RefreshCw size={15} className="text-indigo-500" /> Aylık Net Kâr / Zarar Devri</h2>
          <p className="text-xs text-slate-400 mb-4">Muhasebenizden gelen aylık net sonucu kasaya aktarın.</p>
          <div className="space-y-2 max-h-56 overflow-y-auto pr-1">
            {aylikTrend.length === 0 && <p className="text-xs text-slate-400 text-center py-6">Muhasebe kaydı yok.</p>}
            {aylikTrend.map(m => {
              const yapildi = devirAylar.has(m.ym);
              return (
                <div key={m.ym} className={`flex items-center justify-between p-3 rounded-xl border ${yapildi ? 'bg-slate-50 dark:bg-slate-800/40 border-slate-100 dark:border-slate-700' : 'border-slate-200 dark:border-slate-600 hover:border-indigo-300 dark:hover:border-indigo-600'}`}>
                  <div><span className="text-xs font-semibold">{m.ay}</span><div className="flex gap-3 mt-0.5"><span className="text-xs text-emerald-600">Gelir: {formatTL(m.gelir)}</span><span className="text-xs text-rose-500">Gider: {formatTL(m.gider)}</span></div></div>
                  <div className="flex items-center gap-3">
                    <span className={`text-sm font-bold ${m.kar >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-500'}`}>{m.kar >= 0 ? '+' : ''}{formatTL(m.kar)}</span>
                    {yapildi ? (<span className="flex items-center gap-1 text-xs text-emerald-600 dark:text-emerald-400 font-medium"><CheckCircle2 size={13} /> Aktarıldı</span>) : (<button onClick={() => handleDevir(m.ym, m.kar)} className="text-xs px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-medium whitespace-nowrap">Kasaya Aktar</button>)}
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      </div>
      <Card className="overflow-hidden p-0">
        <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
          <h2 className="text-sm font-semibold flex items-center gap-2"><Filter size={14} className="text-slate-400" /> Kasa Hareketleri <Badge color="violet">{kasaHareketleri.length} kayıt</Badge></h2>
          <p className="text-xs text-slate-400 italic hidden sm:block">Kayıtsız işlem yoktur.</p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-slate-500 dark:text-slate-400">{['Tarih', 'Açıklama', 'Tür', 'Tutar', 'Bakiye', ''].map(h => <th key={h} className="px-4 py-3 text-left font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
            <tbody>
              {sirali.length === 0 && <tr><td colSpan={6} className="px-4 py-8 text-center text-slate-400">Henüz kayıt yok.</td></tr>}
              {sirali.map((h, i) => {
                const gir = h.tur === 'Giriş' || h.tur === 'Kar Devri';
                return (<tr key={h.id} className={`border-t border-slate-100 dark:border-slate-700/60 hover:bg-slate-50 dark:hover:bg-slate-800/40 ${i % 2 ? 'bg-slate-50/50 dark:bg-slate-800/20' : ''}`}>
                  <td className="px-4 py-3 text-slate-500 whitespace-nowrap">{h.tarih}</td>
                  <td className="px-4 py-3 max-w-[240px] truncate" title={h.aciklama}>{h.aciklama}</td>
                  <td className="px-4 py-3"><Badge color={gir ? 'green' : 'red'}>{h.tur}</Badge></td>
                  <td className={`px-4 py-3 font-semibold ${gir ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-500'}`}>{gir ? '+' : '−'}{formatTL(h.tutar)}</td>
                  <td className={`px-4 py-3 font-bold ${h.bakiye >= 0 ? 'text-slate-700 dark:text-slate-200' : 'text-rose-500'}`}>{formatTL(h.bakiye)}</td>
                  <td className="px-4 py-3"><button onClick={() => onKasaDelete(h.id)} className="p-1 rounded text-slate-300 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"><X size={12} /></button></td>
                </tr>);
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── KDV Beyannamesi Sekmesi ──────────────────────────────────
function KdvTab({ transactions }) {
  const aylik = useMemo(() => {
    const map = {};
    transactions.forEach(t => {
      const ym = t.tarih.substring(0, 7);
      if (!map[ym]) map[ym] = { ym, ay: ayEtiketi(ym), alinan: 0, indirilen: 0, islem: 0 };
      map[ym].islem++;
      if (t.tur === 'Gelir') map[ym].alinan += t.kdvTutar;
      else map[ym].indirilen += t.kdvTutar;
    });
    return Object.values(map).sort((a, b) => b.ym.localeCompare(a.ym))
      .map(m => ({ ...m, odenecek: m.alinan - m.indirilen }));
  }, [transactions]);

  const tot = { alinan: aylik.reduce((s, m) => s + m.alinan, 0), indirilen: aylik.reduce((s, m) => s + m.indirilen, 0) };
  tot.odenecek = tot.alinan - tot.indirilen;

  function exportKdv() {
    const h = ['Dönem', 'İşlem', 'Alınan KDV (₺)', 'İndirilecek KDV (₺)', 'Ödenecek KDV (₺)', 'Durum'];
    const d = aylik.map(m => [m.ay, m.islem, m.alinan, m.indirilen, Math.max(0, m.odenecek), m.odenecek >= 0 ? 'Ödeme Var' : 'Devir Alacak']);
    d.push([], ['TOPLAM', aylik.reduce((s, m) => s + m.islem, 0), tot.alinan, tot.indirilen, Math.max(0, tot.odenecek), '']);
    const ws = XLSX.utils.aoa_to_sheet([h, ...d]);
    ws['!cols'] = [10, 8, 16, 18, 16, 14].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'KDV Beyan');
    XLSX.writeFile(wb, `kdv_beyanname_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <Card><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Alınan KDV (Toplam)</p><p className="text-2xl font-bold text-emerald-600 dark:text-emerald-400 mt-1">{formatTL(tot.alinan)}</p><p className="text-xs text-slate-400 mt-1">Gelir işlemlerinden</p></Card>
        <Card><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">İndirilecek KDV (Toplam)</p><p className="text-2xl font-bold text-indigo-600 dark:text-indigo-400 mt-1">{formatTL(tot.indirilen)}</p><p className="text-xs text-slate-400 mt-1">Gider işlemlerinden</p></Card>
        <Card className={tot.odenecek > 0 ? 'border-amber-400 dark:border-amber-600' : 'border-emerald-400 dark:border-emerald-600'}>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Net Ödenecek KDV</p>
          <p className={`text-2xl font-bold mt-1 ${tot.odenecek >= 0 ? 'text-amber-600 dark:text-amber-400' : 'text-emerald-600 dark:text-emerald-400'}`}>{formatTL(Math.abs(tot.odenecek))}</p>
          <p className="text-xs text-slate-400 mt-1">{tot.odenecek >= 0 ? 'Ödeme yapılacak' : 'Devredecek alacak'}</p>
        </Card>
      </div>
      <Card className="overflow-hidden p-0">
        <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
          <h2 className="text-sm font-semibold flex items-center gap-2"><Receipt size={14} className="text-amber-500" /> Aylık KDV Dökümü <span className="text-xs font-normal text-slate-400">— 3065 Sayılı KDV Kanunu</span></h2>
          <button onClick={exportKdv} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium"><FileSpreadsheet size={12} /> Excel'e Aktar</button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-slate-500 dark:text-slate-400">{['Dönem', 'İşlem', 'Alınan KDV', 'İndirilecek KDV', 'Ödenecek KDV', 'Durum'].map(h => <th key={h} className="px-4 py-3 text-left font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
            <tbody>
              {aylik.map((m, i) => (
                <tr key={m.ym} className={`border-t border-slate-100 dark:border-slate-700/60 hover:bg-slate-50 dark:hover:bg-slate-800/40 ${i % 2 ? 'bg-slate-50/50 dark:bg-slate-800/20' : ''}`}>
                  <td className="px-4 py-3 font-medium">{m.ay}</td>
                  <td className="px-4 py-3 text-slate-500">{m.islem}</td>
                  <td className="px-4 py-3 text-emerald-600 dark:text-emerald-400 font-medium">{formatTL(m.alinan)}</td>
                  <td className="px-4 py-3 text-indigo-600 dark:text-indigo-400 font-medium">{formatTL(m.indirilen)}</td>
                  <td className={`px-4 py-3 font-bold ${m.odenecek >= 0 ? 'text-amber-600 dark:text-amber-400' : 'text-emerald-600 dark:text-emerald-400'}`}>{formatTL(Math.abs(m.odenecek))}</td>
                  <td className="px-4 py-3"><Badge color={m.odenecek >= 0 ? 'amber' : 'green'}>{m.odenecek >= 0 ? 'Ödenecek' : 'Alacak'}</Badge></td>
                </tr>
              ))}
              <tr className="border-t-2 border-slate-300 dark:border-slate-600 bg-slate-50 dark:bg-slate-800/60 font-semibold">
                <td className="px-4 py-3">TOPLAM</td>
                <td className="px-4 py-3 text-slate-500">{aylik.reduce((s, m) => s + m.islem, 0)}</td>
                <td className="px-4 py-3 text-emerald-600 dark:text-emerald-400">{formatTL(tot.alinan)}</td>
                <td className="px-4 py-3 text-indigo-600 dark:text-indigo-400">{formatTL(tot.indirilen)}</td>
                <td className={`px-4 py-3 ${tot.odenecek >= 0 ? 'text-amber-600 dark:text-amber-400' : 'text-emerald-600 dark:text-emerald-400'}`}>{formatTL(Math.abs(tot.odenecek))}</td>
                <td className="px-4 py-3"><Badge color={tot.odenecek >= 0 ? 'amber' : 'green'}>{tot.odenecek >= 0 ? 'Ödenecek' : 'Alacak'}</Badge></td>
              </tr>
              {aylik.length === 0 && <tr><td colSpan={6} className="px-4 py-8 text-center text-slate-400">Kayıt bulunamadı.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── Proforma Form Modal ──────────────────────────────────────
function ProformaFormModal({ onClose, onSave, proformalar, editData = null, musteriler = [] }) {
  const today = new Date().toISOString().split('T')[0];
  const plus30 = new Date(Date.now() + 30 * 86400000).toISOString().split('T')[0];
  const [form, setForm] = useState(editData ? { ...editData, kalemler: editData.kalemler.map(k => ({ ...k })) } : {
    faturaNo: nextFaturaNo(proformalar), tarih: today, gecerlilikTarihi: plus30,
    musteri: { ad: '', adres: '', vergiNo: '', vergiDairesi: '' },
    kalemler: [{ id: Date.now(), aciklama: '', miktar: 1, birim: 'Adet', birimFiyat: 0, kdvOrani: 20 }],
    notlar: '',
  });
  const kalemHesap = form.kalemler.map(calcKalem);
  const toplamNet = kalemHesap.reduce((s, k) => s + k.netTutar, 0);
  const toplamKdv = kalemHesap.reduce((s, k) => s + k.kdvTutar, 0);
  const ic = 'w-full text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-1.5 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';
  const setM = (field, val) => setForm(f => ({ ...f, musteri: { ...f.musteri, [field]: val } }));
  const addKalem = () => setForm(f => ({ ...f, kalemler: [...f.kalemler, { id: Date.now(), aciklama: '', miktar: 1, birim: 'Adet', birimFiyat: 0, kdvOrani: 20 }] }));
  const removeKalem = id => setForm(f => ({ ...f, kalemler: f.kalemler.filter(k => k.id !== id) }));
  const updateKalem = (id, field, val) => setForm(f => ({ ...f, kalemler: f.kalemler.map(k => k.id === id ? { ...k, [field]: val } : k) }));
  function handleSubmit(e) {
    e.preventDefault();
    if (!form.musteri.ad.trim()) return;
    onSave({ id: editData ? editData.id : Date.now(), durum: editData?.durum ?? 'Taslak', aktarildi: editData?.aktarildi ?? false, ...form });
    onClose();
  }
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
      <div className="w-full max-w-3xl max-h-[95vh] flex flex-col rounded-2xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-2xl">
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 dark:border-slate-700">
          <h2 className="text-sm font-semibold flex items-center gap-2"><FileText size={15} className="text-indigo-500" /> {editData ? 'Fatura Düzenle' : 'Yeni Proforma Fatura'}</h2>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700"><X size={14} /></button>
        </div>
        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-5 space-y-4">
          <div className="grid grid-cols-3 gap-3">
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Fatura No</label><input value={form.faturaNo} onChange={e => setForm(f => ({ ...f, faturaNo: e.target.value }))} className={ic} /></div>
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Düzenleme Tarihi</label><input type="date" value={form.tarih} onChange={e => setForm(f => ({ ...f, tarih: e.target.value }))} className={ic} /></div>
            <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Geçerlilik Tarihi</label><input type="date" value={form.gecerlilikTarihi} onChange={e => setForm(f => ({ ...f, gecerlilikTarihi: e.target.value }))} className={ic} /></div>
          </div>
          <div className="rounded-xl border border-slate-200 dark:border-slate-600 p-3 space-y-3">
            <div className="flex items-center justify-between">
              <p className="text-xs font-semibold text-slate-600 dark:text-slate-300">Müşteri Bilgileri</p>
              {musteriler.length > 0 && (
                <select onChange={e => { const m = musteriler.find(x => x.id === parseInt(e.target.value)); if (m) setForm(f => ({ ...f, musteri: { ad: m.ad, vergiNo: m.vergiNo || '', vergiDairesi: m.vergiDairesi || '', adres: m.adres || '' } })); }} className="text-xs border border-indigo-200 dark:border-indigo-600 rounded-lg px-2 py-1 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-indigo-600 dark:text-indigo-400">
                  <option value="">Kayıtlı müşteri seç…</option>
                  {musteriler.map(m => <option key={m.id} value={m.id}>{m.ad}</option>)}
                </select>
              )}
            </div>
            <div className="grid grid-cols-3 gap-3">
              <div><label className="text-xs text-slate-500 mb-1 block">Ad / Unvan <span className="text-rose-400">*</span></label><input required value={form.musteri.ad} onChange={e => setM('ad', e.target.value)} placeholder="Müşteri adı" className={ic} /></div>
              <div><label className="text-xs text-slate-500 mb-1 block">Vergi No</label><input value={form.musteri.vergiNo} onChange={e => setM('vergiNo', e.target.value)} placeholder="1234567890" className={ic} /></div>
              <div><label className="text-xs text-slate-500 mb-1 block">Vergi Dairesi</label><input value={form.musteri.vergiDairesi} onChange={e => setM('vergiDairesi', e.target.value)} placeholder="Kadıköy VD" className={ic} /></div>
            </div>
            <div><label className="text-xs text-slate-500 mb-1 block">Adres</label><input value={form.musteri.adres} onChange={e => setM('adres', e.target.value)} placeholder="Müşteri adresi (opsiyonel)" className={ic} /></div>
          </div>
          <div>
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-semibold text-slate-600 dark:text-slate-300">Kalemler</p>
              <button type="button" onClick={addKalem} className="text-xs px-2 py-1 rounded-lg bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-900/50 font-medium flex items-center gap-1"><PlusCircle size={11} /> Satır Ekle</button>
            </div>
            <div className="rounded-xl border border-slate-200 dark:border-slate-600 overflow-hidden">
              <table className="w-full text-xs">
                <thead><tr className="bg-slate-50 dark:bg-slate-800 text-slate-500">{['Açıklama', 'Miktar', 'Birim', 'Birim Fiyat', 'KDV %', 'Toplam', ''].map(h => <th key={h} className="px-2 py-2 text-left font-medium">{h}</th>)}</tr></thead>
                <tbody>
                  {form.kalemler.map(k => {
                    const kh = calcKalem(k);
                    return (
                      <tr key={k.id} className="border-t border-slate-100 dark:border-slate-700">
                        <td className="px-2 py-1.5"><input value={k.aciklama} onChange={e => updateKalem(k.id, 'aciklama', e.target.value)} placeholder="Ürün / Hizmet" className={ic} /></td>
                        <td className="px-2 py-1.5 w-16"><input type="number" min="0" step="0.01" value={k.miktar} onChange={e => updateKalem(k.id, 'miktar', parseFloat(e.target.value) || 0)} className={ic} /></td>
                        <td className="px-2 py-1.5 w-20"><select value={k.birim} onChange={e => updateKalem(k.id, 'birim', e.target.value)} className={ic}>{BIRIMLER.map(b => <option key={b}>{b}</option>)}</select></td>
                        <td className="px-2 py-1.5 w-24"><input type="number" min="0" step="0.01" value={k.birimFiyat} onChange={e => updateKalem(k.id, 'birimFiyat', parseFloat(e.target.value) || 0)} className={ic} /></td>
                        <td className="px-2 py-1.5 w-20"><select value={k.kdvOrani} onChange={e => updateKalem(k.id, 'kdvOrani', parseInt(e.target.value))} className={ic}>{KDV_ORANLARI.map(o => <option key={o} value={o}>%{o}</option>)}</select></td>
                        <td className="px-3 py-1.5 font-medium whitespace-nowrap">{formatTL(kh.toplam)}</td>
                        <td className="px-2 py-1.5"><button type="button" onClick={() => removeKalem(k.id)} disabled={form.kalemler.length === 1} className="p-1 rounded text-slate-300 hover:text-rose-500 disabled:opacity-20 transition-colors"><X size={11} /></button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="mt-2 flex justify-end">
              <div className="rounded-xl bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 p-3 text-xs space-y-1 min-w-[200px]">
                <div className="flex justify-between gap-6"><span className="text-slate-500">Ara Toplam</span><span className="font-medium">{formatTL(toplamNet)}</span></div>
                <div className="flex justify-between gap-6"><span className="text-slate-500">KDV</span><span className="font-medium text-amber-600 dark:text-amber-400">{formatTL(toplamKdv)}</span></div>
                <div className="flex justify-between gap-6 border-t border-slate-200 dark:border-slate-600 pt-1"><span className="font-semibold">Genel Toplam</span><span className="font-bold text-indigo-600 dark:text-indigo-400">{formatTL(toplamNet + toplamKdv)}</span></div>
              </div>
            </div>
          </div>
          <div><label className="text-xs font-medium text-slate-500 mb-1.5 block">Notlar / Ödeme Koşulları</label><textarea value={form.notlar} onChange={e => setForm(f => ({ ...f, notlar: e.target.value }))} rows={2} placeholder="Opsiyonel notlar..." className={`${ic} resize-none`} /></div>
          <div className="flex gap-2 pt-1">
            <button type="button" onClick={onClose} className="flex-1 text-xs py-2.5 rounded-xl border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">İptal</button>
            <button type="submit" className="flex-1 text-xs py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium">Kaydet</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ── Proforma Print Modal ─────────────────────────────────────
function ProformaPrintModal({ fatura, onClose, logo = null }) {
  const kalemler = fatura.kalemler.map(calcKalem);
  const toplamNet = kalemler.reduce((s, k) => s + k.netTutar, 0);
  const toplamKdv = kalemler.reduce((s, k) => s + k.kdvTutar, 0);
  useEffect(() => {
    const s = document.createElement('style');
    s.id = '__prf_print__';
    s.textContent = `@media print { body * { visibility: hidden !important; } #prf-print-root, #prf-print-root * { visibility: visible !important; } #prf-print-root { position: fixed; inset: 0; background: white; z-index: 99999; padding: 32px; overflow: auto; } }`;
    document.head.appendChild(s);
    return () => { document.getElementById('__prf_print__')?.remove(); };
  }, []);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
      <div className="w-full max-w-2xl max-h-[95vh] flex flex-col rounded-2xl bg-white shadow-2xl">
        <div className="flex items-center justify-between px-5 py-3 border-b border-slate-200">
          <h2 className="text-sm font-semibold text-slate-800">Fatura Önizleme</h2>
          <div className="flex gap-2">
            <button onClick={() => exportProformaExcel(fatura)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 font-medium text-slate-700"><FileSpreadsheet size={12} /> Excel</button>
            <button onClick={() => window.print()} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-medium">Yazdır / PDF</button>
            <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100"><X size={14} className="text-slate-600" /></button>
          </div>
        </div>
        <div id="prf-print-root" className="flex-1 overflow-y-auto p-8 text-slate-800">
          <div className="flex items-start justify-between mb-8">
            <div>
              {logo && <img src={logo} alt="Logo" className="h-14 object-contain mb-3" />}
              <h1 className="text-2xl font-bold text-indigo-700">PROFORMA FATURA</h1><p className="text-sm text-slate-500 mt-1">No: <span className="font-semibold text-slate-700">{fatura.faturaNo}</span></p>
            </div>
            <div className="text-right text-xs text-slate-500 space-y-0.5"><p>Tarih: <span className="font-medium text-slate-700">{fatura.tarih}</span></p><p>Geçerlilik: <span className="font-medium text-slate-700">{fatura.gecerlilikTarihi}</span></p></div>
          </div>
          <div className="mb-6 p-4 rounded-xl bg-slate-50 border border-slate-200">
            <p className="text-xs font-semibold uppercase tracking-wide text-slate-400 mb-2">Müşteri</p>
            <p className="font-semibold">{fatura.musteri.ad}</p>
            {fatura.musteri.vergiNo && <p className="text-xs text-slate-500 mt-0.5">VKN: {fatura.musteri.vergiNo}{fatura.musteri.vergiDairesi && ` — ${fatura.musteri.vergiDairesi} V.D.`}</p>}
            {fatura.musteri.adres && <p className="text-xs text-slate-500 mt-0.5">{fatura.musteri.adres}</p>}
          </div>
          <table className="w-full text-sm mb-6 border-collapse">
            <thead><tr className="border-b-2 border-slate-300">{['#', 'Açıklama', 'Miktar', 'Birim Fiyat', 'KDV %', 'Toplam'].map(h => <th key={h} className="text-left py-2 text-xs font-semibold text-slate-600">{h}</th>)}</tr></thead>
            <tbody>
              {kalemler.map((k, i) => (
                <tr key={k.id} className="border-b border-slate-100">
                  <td className="py-2 text-xs text-slate-400 pr-3">{i + 1}</td>
                  <td className="py-2 text-xs">{k.aciklama} <span className="text-slate-400">({k.birim})</span></td>
                  <td className="py-2 text-xs text-right pr-4">{k.miktar}</td>
                  <td className="py-2 text-xs text-right pr-4">{formatTL(k.birimFiyat)}</td>
                  <td className="py-2 text-xs text-right pr-4">%{k.kdvOrani}</td>
                  <td className="py-2 text-xs text-right font-medium">{formatTL(k.toplam)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="flex justify-end mb-6">
            <div className="space-y-1 min-w-[200px] text-sm">
              <div className="flex justify-between gap-8"><span className="text-slate-500">Ara Toplam</span><span>{formatTL(toplamNet)}</span></div>
              <div className="flex justify-between gap-8"><span className="text-slate-500">KDV</span><span>{formatTL(toplamKdv)}</span></div>
              <div className="flex justify-between gap-8 border-t border-slate-200 pt-1 font-bold"><span>GENEL TOPLAM</span><span className="text-indigo-700">{formatTL(toplamNet + toplamKdv)}</span></div>
            </div>
          </div>
          {fatura.notlar && <div className="p-3 rounded-xl border border-slate-200 bg-slate-50 text-xs text-slate-600"><p className="font-semibold mb-1">Notlar</p><p>{fatura.notlar}</p></div>}
          <p className="text-xs text-slate-400 mt-6 text-center">Bu belge proforma fatura niteliğindedir. Resmi fatura değildir.</p>
        </div>
      </div>
    </div>
  );
}

// ── Proforma Sekmesi ─────────────────────────────────────────
function ProformaTab({ proformalar, onAdd, onUpdate, onDelete, onAktar, musteriler = [], logo = null }) {
  const [showForm, setShowForm] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [printItem, setPrintItem] = useState(null);
  function openEdit(p) { setEditItem(p); setShowForm(true); }
  function closeForm() { setShowForm(false); setEditItem(null); }
  return (
    <div className="space-y-6">
      {showForm && <ProformaFormModal onClose={closeForm} onSave={p => { editItem ? onUpdate(p) : onAdd(p); }} proformalar={proformalar} editData={editItem} musteriler={musteriler} />}
      {printItem && <ProformaPrintModal fatura={printItem} onClose={() => setPrintItem(null)} logo={logo} />}
      <div className="flex items-center justify-between">
        <div><h2 className="text-base font-semibold">Proforma Faturalar</h2><p className="text-xs text-slate-400 mt-0.5">{proformalar.length} fatura</p></div>
        <button onClick={() => { setEditItem(null); setShowForm(true); }} className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white"><PlusCircle size={13} /> Yeni Fatura</button>
      </div>
      {proformalar.length === 0 ? (
        <Card className="py-16 text-center">
          <FileText size={32} className="mx-auto text-slate-300 dark:text-slate-600 mb-3" />
          <p className="text-sm text-slate-400">Henüz proforma fatura oluşturulmadı.</p>
          <button onClick={() => setShowForm(true)} className="mt-3 text-xs text-indigo-600 dark:text-indigo-400 hover:underline">İlk faturayı oluştur</button>
        </Card>
      ) : (
        <Card className="overflow-hidden p-0">
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-slate-500">{['Fatura No', 'Tarih', 'Geçerlilik', 'Müşteri', 'Tutar', 'Durum', ''].map(h => <th key={h} className="px-4 py-3 text-left font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
              <tbody>
                {[...proformalar].sort((a, b) => b.tarih.localeCompare(a.tarih)).map((p, i) => {
                  const tutar = p.kalemler.map(calcKalem).reduce((s, k) => s + k.toplam, 0);
                  const durumRenk = { Taslak: 'blue', Onaylandı: 'green', İptal: 'red' };
                  return (
                    <tr key={p.id} className={`border-t border-slate-100 dark:border-slate-700/60 hover:bg-slate-50 dark:hover:bg-slate-800/40 ${i % 2 ? 'bg-slate-50/50 dark:bg-slate-800/20' : ''}`}>
                      <td className="px-4 py-3 font-mono font-medium">{p.faturaNo}</td>
                      <td className="px-4 py-3 text-slate-500">{p.tarih}</td>
                      <td className="px-4 py-3 text-slate-500">{p.gecerlilikTarihi}</td>
                      <td className="px-4 py-3 font-medium max-w-[160px] truncate" title={p.musteri.ad}>{p.musteri.ad}</td>
                      <td className="px-4 py-3 font-semibold">{formatTL(tutar)}</td>
                      <td className="px-4 py-3">
                        <select value={p.durum} onChange={e => onUpdate({ ...p, durum: e.target.value })} className="text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-1 bg-white dark:bg-slate-700 focus:outline-none">
                          {['Taslak', 'Onaylandı', 'İptal'].map(d => <option key={d}>{d}</option>)}
                        </select>
                      </td>
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-1">
                          <button onClick={() => setPrintItem(p)} title="Önizle / PDF" className="p-1 rounded text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors"><Receipt size={12} /></button>
                          <button onClick={() => openEdit(p)} title="Düzenle" className="p-1 rounded text-slate-400 hover:text-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"><ClipboardList size={12} /></button>
                          {p.durum === 'Onaylandı' && !p.aktarildi && (
                            <button onClick={() => onAktar(p)} title="Muhasebeye Aktar (Gelir kaydı oluştur)" className="p-1 rounded text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 transition-colors"><CheckCircle2 size={12} /></button>
                          )}
                          {p.aktarildi && <CheckCircle2 size={12} className="text-emerald-500" title="Muhasebeye aktarıldı" />}
                          <button onClick={() => onDelete(p.id)} className="p-1 rounded text-slate-300 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"><X size={12} /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}

// ── Ayarlar Sekmesi ──────────────────────────────────────────
function AyarlarTab({ personeller, onPersonelEkle, onPersonelSil, transactions, kasaHareketleri, onResetAll, butceler, onButceChange, musteriler, onMusteriEkle, onMusteriSil, logo, onLogoChange }) {
  const [yeni, setYeni] = useState('');
  const [resetConfirm, setResetConfirm] = useState(false);
  const [musteriForm, setMusteriForm] = useState({ ad: '', vergiNo: '', vergiDairesi: '', adres: '', email: '', telefon: '' });
  const logoInputRef = useRef(null);
  function handleMusteriEkle(e) {
    e.preventDefault();
    if (!musteriForm.ad.trim()) return;
    onMusteriEkle({ id: Date.now(), ...musteriForm, ad: musteriForm.ad.trim() });
    setMusteriForm({ ad: '', vergiNo: '', vergiDairesi: '', adres: '', email: '', telefon: '' });
  }
  function handleLogoUpload(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => onLogoChange(evt.target.result);
    reader.readAsDataURL(file);
    e.target.value = '';
  }
  function handleEkle(e) {
    e.preventDefault();
    const isim = yeni.trim();
    if (!isim || personeller.map(p => p.toLowerCase()).includes(isim.toLowerCase())) return;
    onPersonelEkle(isim);
    setYeni('');
  }
  const ic = 'text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500';
  const micInput = `flex-1 text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500`;
  return (
    <div className="space-y-6 max-w-2xl">

      {/* Logo */}
      <Card>
        <h2 className="text-sm font-semibold mb-3 flex items-center gap-2"><Receipt size={15} className="text-indigo-500" /> Firma Logosu</h2>
        <input ref={logoInputRef} type="file" accept="image/*" className="hidden" onChange={handleLogoUpload} />
        {logo ? (
          <div className="flex items-center gap-4">
            <img src={logo} alt="Logo" className="h-16 object-contain rounded-lg border border-slate-200 dark:border-slate-600 p-2 bg-white" />
            <div className="space-y-2">
              <p className="text-xs text-slate-500">Proforma fatura önizlemesinde görünür.</p>
              <div className="flex gap-2">
                <button onClick={() => logoInputRef.current?.click()} className="text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">Değiştir</button>
                <button onClick={() => onLogoChange('')} className="text-xs px-3 py-1.5 rounded-lg border border-rose-200 dark:border-rose-800 text-rose-600 dark:text-rose-400 hover:bg-rose-50 dark:hover:bg-rose-900/20 font-medium">Kaldır</button>
              </div>
            </div>
          </div>
        ) : (
          <button onClick={() => logoInputRef.current?.click()} className="flex items-center gap-2 text-xs px-4 py-2.5 rounded-xl border-2 border-dashed border-slate-200 dark:border-slate-600 hover:border-indigo-400 dark:hover:border-indigo-500 text-slate-500 hover:text-indigo-600 dark:hover:text-indigo-400 transition-colors font-medium w-full justify-center">
            <Upload size={14} /> Logo Yükle (PNG, JPG, SVG)
          </button>
        )}
      </Card>

      {/* Müşteri Rehberi */}
      <Card>
        <h2 className="text-sm font-semibold mb-4 flex items-center gap-2"><ClipboardList size={15} className="text-indigo-500" /> Müşteri Rehberi</h2>
        <form onSubmit={handleMusteriEkle} className="space-y-2 mb-5 p-3 rounded-xl bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600">
          <p className="text-xs font-medium text-slate-500 mb-2">Yeni Müşteri Ekle</p>
          <div className="grid grid-cols-2 gap-2">
            <input value={musteriForm.ad} onChange={e => setMusteriForm(f => ({ ...f, ad: e.target.value }))} placeholder="Şirket / Müşteri adı *" required className={micInput} />
            <input value={musteriForm.vergiNo} onChange={e => setMusteriForm(f => ({ ...f, vergiNo: e.target.value }))} placeholder="Vergi No" className={micInput} />
            <input value={musteriForm.vergiDairesi} onChange={e => setMusteriForm(f => ({ ...f, vergiDairesi: e.target.value }))} placeholder="Vergi Dairesi" className={micInput} />
            <input value={musteriForm.telefon} onChange={e => setMusteriForm(f => ({ ...f, telefon: e.target.value }))} placeholder="Telefon" className={micInput} />
          </div>
          <input value={musteriForm.adres} onChange={e => setMusteriForm(f => ({ ...f, adres: e.target.value }))} placeholder="Adres" className={`w-full ${micInput}`} />
          <div className="flex justify-end">
            <button type="submit" disabled={!musteriForm.ad.trim()} className="text-xs px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white font-medium flex items-center gap-1.5"><UserPlus size={12} /> Kaydet</button>
          </div>
        </form>
        <div className="space-y-2">
          {musteriler.length === 0 && <p className="text-xs text-slate-400 text-center py-4">Henüz müşteri eklenmemiş.</p>}
          {musteriler.map(m => (
            <div key={m.id} className="flex items-center justify-between p-3 rounded-xl border border-slate-200 dark:border-slate-600">
              <div>
                <span className="text-sm font-medium">{m.ad}</span>
                {m.vergiNo && <span className="ml-2 text-xs text-slate-400">VKN: {m.vergiNo}</span>}
                {m.adres && <p className="text-xs text-slate-400 mt-0.5">{m.adres}</p>}
              </div>
              <button onClick={() => onMusteriSil(m.id)} className="p-1.5 rounded-lg text-slate-400 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"><Trash2 size={13} /></button>
            </div>
          ))}
        </div>
      </Card>

      {/* Personel Yönetimi */}
      <Card>
        <h2 className="text-sm font-semibold mb-4 flex items-center gap-2"><UserPlus size={15} className="text-indigo-500" /> Personel Yönetimi</h2>
        <form onSubmit={handleEkle} className="flex gap-2 mb-5">
          <input value={yeni} onChange={e => setYeni(e.target.value)} placeholder="Yeni personel adı..." className={`flex-1 ${ic}`} />
          <button type="submit" disabled={!yeni.trim()} className="text-xs px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white font-medium flex items-center gap-1.5"><UserPlus size={12} /> Ekle</button>
        </form>
        <div className="space-y-2">
          {personeller.map(p => {
            const islem = transactions.filter(t => t.personel === p).length;
            return (
              <div key={p} className="flex items-center justify-between p-3 rounded-xl border border-slate-200 dark:border-slate-600">
                <div><span className="text-sm font-medium">{p}</span><span className="ml-2 text-xs text-slate-400">{islem} işlem kaydı</span></div>
                <button onClick={() => onPersonelSil(p)} disabled={islem > 0} title={islem > 0 ? 'İşlem kaydı olan personel silinemez' : 'Personeli sil'}
                  className="p-1.5 rounded-lg text-slate-400 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 disabled:opacity-30 disabled:cursor-not-allowed transition-colors">
                  <Trash2 size={13} />
                </button>
              </div>
            );
          })}
          {personeller.length === 0 && <p className="text-xs text-slate-400 text-center py-4">Personel eklenmemiş.</p>}
        </div>
      </Card>

      {/* Bütçe Limitleri */}
      <Card>
        <h2 className="text-sm font-semibold mb-1 flex items-center gap-2"><Target size={15} className="text-indigo-500" /> Aylık Bütçe Limitleri</h2>
        <p className="text-xs text-slate-400 mb-4">Kategori başına aylık harcama limiti (₺). 0 = limit yok.</p>
        <div className="space-y-3">
          {KATEGORILER.map(k => (
            <div key={k} className="flex items-center gap-3">
              <span className="text-xs font-medium w-28 shrink-0">{k}</span>
              <input type="number" min="0" step="100" value={butceler[k] || 0}
                onChange={e => onButceChange(k, Math.max(0, parseInt(e.target.value) || 0))}
                className={`flex-1 ${ic}`} placeholder="0" />
              <span className="text-xs text-slate-400 shrink-0">₺/ay</span>
            </div>
          ))}
        </div>
      </Card>

      {/* Veri Dışa Aktarma */}
      <Card>
        <h2 className="text-sm font-semibold mb-4 flex items-center gap-2"><FileSpreadsheet size={15} className="text-emerald-500" /> Veri Dışa Aktarma</h2>
        <div className="space-y-3">
          <div className="flex items-center justify-between p-3 rounded-xl border border-slate-200 dark:border-slate-600">
            <div><p className="text-xs font-medium">Tüm Muhasebe Kayıtları</p><p className="text-xs text-slate-400 mt-0.5">{transactions.length} işlem, Excel formatında</p></div>
            <button onClick={() => exportMuhasebeExcel([...transactions].sort((a, b) => new Date(b.tarih) - new Date(a.tarih)), 'tumu')}
              className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">
              <Download size={12} /> İndir
            </button>
          </div>
          <div className="flex items-center justify-between p-3 rounded-xl border border-slate-200 dark:border-slate-600">
            <div><p className="text-xs font-medium">Tüm Kasa Hareketleri</p><p className="text-xs text-slate-400 mt-0.5">{kasaHareketleri.length} hareket, bakiye sütunuyla</p></div>
            <button onClick={() => exportKasaExcel(kasaHareketleri)}
              className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">
              <Download size={12} /> İndir
            </button>
          </div>
        </div>
      </Card>

      {/* Tehlikeli Alan */}
      <Card className="border-rose-200 dark:border-rose-800">
        <h2 className="text-sm font-semibold mb-1 flex items-center gap-2 text-rose-600 dark:text-rose-400"><AlertTriangle size={15} /> Tehlikeli Alan</h2>
        <p className="text-xs text-slate-500 dark:text-slate-400 mb-4">Bu işlem geri alınamaz. Tüm muhasebe ve kasa verileri silinir, uygulama başlangıç durumuna döner.</p>
        {!resetConfirm ? (
          <button onClick={() => setResetConfirm(true)} className="text-xs px-4 py-2 rounded-lg border border-rose-300 dark:border-rose-700 text-rose-600 dark:text-rose-400 hover:bg-rose-50 dark:hover:bg-rose-900/20 font-medium transition-colors">
            Tüm Verileri Sıfırla
          </button>
        ) : (
          <div className="flex items-center gap-3">
            <span className="text-xs text-rose-600 dark:text-rose-400 font-medium">Emin misiniz? Bu işlem geri alınamaz!</span>
            <button onClick={() => { onResetAll(); setResetConfirm(false); }} className="text-xs px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white font-medium">Evet, Sıfırla</button>
            <button onClick={() => setResetConfirm(false)} className="text-xs px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 font-medium">İptal</button>
          </div>
        )}
      </Card>
    </div>
  );
}

// ── Ana Uygulama ─────────────────────────────────────────────
export default function App() {
  const [isDark, setIsDark] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');

  // localStorage'dan başlat
  const [transactions, setTransactions]     = useState(() => lsGet(LS.tx, MOCK_DATA));
  const [kasaHareketleri, setKasaHareketleri] = useState(() => lsGet(LS.kasa, MOCK_KASA));
  const [personeller, setPersoneller]       = useState(() => lsGet(LS.personeller, DEFAULT_PERSONELLER));
  const [butceler, setButceler]             = useState(() => lsGet(LS.butce, BUTCE_DEFAULT));
  const [proformalar, setProformalar]       = useState(() => lsGet(LS.proforma, []));
  const [musteriler, setMusteriler]         = useState(() => lsGet(LS.musteriler, []));
  const [logo, setLogo]                     = useState(() => lsGet(LS.logo, ''));
  const [sabitGiderler, setSabitGiderler]   = useState(() => lsGet(LS.sabitGiderler, []));

  // localStorage'a kaydet
  useEffect(() => lsSet(LS.tx, transactions),          [transactions]);
  useEffect(() => lsSet(LS.kasa, kasaHareketleri),     [kasaHareketleri]);
  useEffect(() => lsSet(LS.personeller, personeller),  [personeller]);
  useEffect(() => lsSet(LS.butce, butceler),           [butceler]);
  useEffect(() => lsSet(LS.proforma, proformalar),     [proformalar]);
  useEffect(() => lsSet(LS.musteriler, musteriler),    [musteriler]);
  useEffect(() => lsSet(LS.logo, logo),                [logo]);
  useEffect(() => lsSet(LS.sabitGiderler, sabitGiderler), [sabitGiderler]);

  // Chat
  const [chatInput, setChatInput] = useState('');
  const [chatLog, setChatLog] = useState([
    { from: 'bot', text: 'Merhaba! Doğal dille veri girebilirsiniz. Örnek: "Personel 2 bugün 1200 TL pazarlama harcaması yaptı, %20 KDV dahil"' }
  ]);

  // Tablo filtreleri
  const [filterTur,      setFilterTur]      = useState('Tümü');
  const [filterPersonel, setFilterPersonel] = useState('Tümü');
  const [filterKategori, setFilterKategori] = useState('Tümü');

  // Tarih aralığı filtresi
  const [datePreset, setDatePreset] = useState('tumu');
  const [dateFrom,   setDateFrom]   = useState('');
  const [dateTo,     setDateTo]     = useState('');

  // Modal state
  const [showForm,    setShowForm]    = useState(false);
  const [excelPreview, setExcelPreview] = useState(null);
  const fileInputRef = useRef(null);

  // Aktif tarih aralığı
  const dateRange = useMemo(() => calcDateRange(datePreset, dateFrom, dateTo), [datePreset, dateFrom, dateTo]);

  // Tarih filtresine göre transactions
  const dateTx = useMemo(() => {
    const { from, to } = dateRange;
    if (!from && !to) return transactions;
    return transactions.filter(t => (!from || t.tarih >= from) && (!to || t.tarih <= to));
  }, [transactions, dateRange]);

  // Özet (tarih filtresine duyarlı)
  const ozet = useMemo(() => {
    const gelir = dateTx.filter(t => t.tur === 'Gelir');
    const gider = dateTx.filter(t => t.tur === 'Gider');
    return {
      toplamGelir: gelir.reduce((s, t) => s + t.netTutar, 0),
      toplamGider: gider.reduce((s, t) => s + t.netTutar, 0),
      netKar:      gelir.reduce((s, t) => s + t.netTutar, 0) - gider.reduce((s, t) => s + t.netTutar, 0),
      odenecekKdv: gelir.reduce((s, t) => s + t.kdvTutar, 0) - gider.reduce((s, t) => s + t.kdvTutar, 0),
    };
  }, [dateTx]);

  const personelHarcama = useMemo(() => personeller.map(p => ({
    name: p, toplam: dateTx.filter(t => t.tur === 'Gider' && t.personel === p).reduce((s, t) => s + t.toplamTutar, 0),
  })), [dateTx, personeller]);

  const kategoriDagilim = useMemo(() => {
    const map = {};
    dateTx.filter(t => t.tur === 'Gider').forEach(t => { map[t.kategori] = (map[t.kategori] || 0) + t.netTutar; });
    return Object.entries(map).map(([name, value]) => ({ name, value }));
  }, [dateTx]);

  // Trend her zaman tüm veri (tarih filtresi trend grafiğini etkilemez)
  const aylikTrend = useMemo(() => {
    const map = {};
    transactions.forEach(t => {
      const ym = t.tarih.substring(0, 7);
      if (!map[ym]) map[ym] = { ym, gelir: 0, gider: 0 };
      if (t.tur === 'Gelir') map[ym].gelir += t.netTutar; else map[ym].gider += t.netTutar;
    });
    return Object.values(map).sort((a, b) => a.ym.localeCompare(b.ym)).map(m => ({ ...m, ay: ayEtiketi(m.ym), kar: m.gelir - m.gider }));
  }, [transactions]);

  // Filtrelenmiş tablo (tarih + tür + personel + kategori)
  const filteredRows = useMemo(() =>
    dateTx.filter(t => {
      if (filterTur !== 'Tümü' && t.tur !== filterTur) return false;
      if (filterPersonel !== 'Tümü' && t.personel !== filterPersonel) return false;
      if (filterKategori !== 'Tümü' && t.kategori !== filterKategori) return false;
      return true;
    }).sort((a, b) => new Date(b.tarih) - new Date(a.tarih)),
    [dateTx, filterTur, filterPersonel, filterKategori]);

  async function handleChatSubmit(e) {
    e.preventDefault();
    const msg = chatInput.trim();
    if (!msg) return;
    setChatInput('');
    setChatLog(p => [...p, { from: 'user', text: msg }, { from: 'bot', text: '...', loading: true }]);
    try {
      const result = await askGemini(msg, transactions, personeller);
      if (result.type === 'transaction') {
        setTransactions(p => [result.data, ...p]);
        const d = result.data;
        setChatLog(p => [...p.slice(0, -1), { from: 'bot', text: `✅ Eklendi: ${d.tur} — ${formatTL(d.toplamTutar)} (KDV %${d.kdvOrani}) | ${d.kategori}${d.personel ? ` | ${d.personel}` : ''}` }]);
      } else {
        setChatLog(p => [...p.slice(0, -1), { from: 'bot', text: result.text }]);
      }
    } catch (_) {
      const parsed = parseChat(msg, personeller);
      const bot = parsed
        ? { from: 'bot', text: `✅ Eklendi: ${parsed.tur} — ${formatTL(parsed.toplamTutar)} (KDV %${parsed.kdvOrani}) | ${parsed.kategori}${parsed.personel ? ` | ${parsed.personel}` : ''}` }
        : { from: 'bot', text: _.message === 'QUOTA' ? '⚠️ Gemini kota limitine ulaşıldı. Google AI Studio\'dan kotanızı kontrol edin.' : '❌ Gemini bağlanamadı. Lütfen tekrar deneyin.' };
      if (parsed) setTransactions(p => [parsed, ...p]);
      setChatLog(p => [...p.slice(0, -1), bot]);
    }
  }

  function handleExcelUpload(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => { const wb = XLSX.read(evt.target.result, { type: 'binary' }); setExcelPreview(parseExcelRows(wb, personeller)); };
    reader.readAsBinaryString(file);
    e.target.value = '';
  }

  function handleResetAll() {
    setTransactions([]);
    setKasaHareketleri([]);
    setPersoneller(DEFAULT_PERSONELLER);
    setButceler(BUTCE_DEFAULT);
    setProformalar([]);
    lsSet(LS.tx, []); lsSet(LS.kasa, []); lsSet(LS.personeller, DEFAULT_PERSONELLER); lsSet(LS.butce, BUTCE_DEFAULT); lsSet(LS.proforma, []);
  }

  const toggleDark = () => { setIsDark(d => !d); document.documentElement.classList.toggle('dark'); };

  const TABS = [
    { id: 'dashboard', label: 'Dashboard',     icon: <LayoutDashboard size={13} /> },
    { id: 'kasa',      label: 'Kasa & Banka',  icon: <Landmark size={13} /> },
    { id: 'kdv',       label: 'KDV Beyanname',    icon: <FileText size={13} /> },
    { id: 'proforma',  label: 'Proforma Fatura', icon: <Receipt size={13} /> },
    { id: 'ayarlar',   label: 'Ayarlar',          icon: <Settings size={13} /> },
  ];

  return (
    <div className={isDark ? 'dark' : ''}>
      <div className="min-h-screen bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 transition-colors duration-300">

        {showForm && <FormModal onClose={() => setShowForm(false)} onAdd={r => setTransactions(p => [r, ...p])} personeller={personeller} />}
        {excelPreview && <ExcelPreviewModal parsed={excelPreview.parsed} errors={excelPreview.errors} onConfirm={rows => { setTransactions(p => [...rows, ...p]); setExcelPreview(null); }} onClose={() => setExcelPreview(null)} />}
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelUpload} />

        {/* HEADER */}
        <header className="sticky top-0 z-10 bg-white/80 dark:bg-slate-900/80 backdrop-blur border-b border-slate-200 dark:border-slate-700 px-4 py-3 flex items-center justify-between gap-2">
          <div className="flex items-center gap-3 shrink-0">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center"><BarChart2 size={18} className="text-white" /></div>
            <div className="hidden sm:block"><h1 className="text-base font-semibold leading-none">Muhasebe Dashboard</h1><p className="text-xs text-slate-500 dark:text-slate-400">Gelir-Gider & KDV Takip</p></div>
          </div>
          <div className="flex items-center gap-2 flex-wrap justify-end">
            <button onClick={() => setShowForm(true)} className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white transition-colors"><PlusCircle size={13} /> Manuel Giriş</button>
            <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"><Upload size={13} /> Excel'den Aktar</button>
            <button onClick={downloadTemplate} className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"><Download size={13} /> Şablon</button>
            <button onClick={toggleDark} className="p-2 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors">{isDark ? <Sun size={18} /> : <Moon size={18} />}</button>
          </div>
        </header>

        {/* TAB NAV */}
        <div className="sticky top-[57px] z-10 bg-white/90 dark:bg-slate-900/90 backdrop-blur border-b border-slate-200 dark:border-slate-700 px-4">
          <div className="flex max-w-7xl mx-auto">
            {TABS.map(tab => (
              <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                className={`flex items-center gap-1.5 px-4 py-3 text-xs font-medium border-b-2 transition-colors ${activeTab === tab.id ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400' : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'}`}>
                {tab.icon} {tab.label}
              </button>
            ))}
          </div>
        </div>

        <main className="max-w-7xl mx-auto px-4 py-6 space-y-6">

          {/* ── DASHBOARD ── */}
          {activeTab === 'dashboard' && (<>

            {/* TARİH FİLTRESİ */}
            <Card className="py-3 px-4">
              <DateFilterBar
                preset={datePreset} from={dateFrom} to={dateTo}
                onPreset={setDatePreset} onFrom={setDateFrom} onTo={setDateTo}
                count={dateTx.length} total={transactions.length}
              />
            </Card>

            {/* ÖZET KARTLARI */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <Card>
                <div className="flex items-start justify-between">
                  <div><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Toplam Gelir</p><p className="text-2xl font-bold text-emerald-600 dark:text-emerald-400 mt-1">{formatTL(ozet.toplamGelir)}</p></div>
                  <div className="w-10 h-10 rounded-xl bg-emerald-100 dark:bg-emerald-900/40 flex items-center justify-center"><TrendingUp size={20} className="text-emerald-600 dark:text-emerald-400" /></div>
                </div>
              </Card>
              <Card>
                <div className="flex items-start justify-between">
                  <div><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Toplam Gider</p><p className="text-2xl font-bold text-rose-600 dark:text-rose-400 mt-1">{formatTL(ozet.toplamGider)}</p></div>
                  <div className="w-10 h-10 rounded-xl bg-rose-100 dark:bg-rose-900/40 flex items-center justify-center"><TrendingDown size={20} className="text-rose-600 dark:text-rose-400" /></div>
                </div>
              </Card>
              <Card>
                <div className="flex items-start justify-between">
                  <div><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Net Kar/Zarar</p><p className={`text-2xl font-bold mt-1 ${ozet.netKar >= 0 ? 'text-indigo-600 dark:text-indigo-400' : 'text-rose-600'}`}>{formatTL(ozet.netKar)}</p></div>
                  <div className="w-10 h-10 rounded-xl bg-indigo-100 dark:bg-indigo-900/40 flex items-center justify-center"><Wallet size={20} className="text-indigo-600 dark:text-indigo-400" /></div>
                </div>
              </Card>
              <Card className={ozet.odenecekKdv > 0 ? 'border-amber-400 dark:border-amber-600' : ''}>
                <div className="flex items-start justify-between">
                  <div><p className="text-xs text-slate-500 dark:text-slate-400 font-medium uppercase tracking-wide">Ödenecek KDV</p><p className="text-2xl font-bold text-amber-600 dark:text-amber-400 mt-1">{formatTL(Math.max(0, ozet.odenecekKdv))}</p><p className="text-xs text-slate-400 mt-1">3065 Sayılı KDV Kanunu</p></div>
                  <div className="w-10 h-10 rounded-xl bg-amber-100 dark:bg-amber-900/40 flex items-center justify-center"><Receipt size={20} className="text-amber-600 dark:text-amber-400" /></div>
                </div>
              </Card>
            </div>

            {/* BÜTÇE TAKİBİ */}
            {KATEGORILER.some(k => butceler[k] > 0) && (() => {
              const now = new Date(); const ym = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
              const ayGider = {};
              dateTx.filter(t => t.tur === 'Gider' && t.tarih.startsWith(ym)).forEach(t => { ayGider[t.kategori] = (ayGider[t.kategori] || 0) + t.netTutar; });
              const rows = KATEGORILER.filter(k => butceler[k] > 0).map(k => ({ k, harcanan: ayGider[k] || 0, limit: butceler[k], pct: Math.min(100, Math.round(((ayGider[k] || 0) / butceler[k]) * 100)) }));
              return (
                <Card>
                  <h2 className="text-sm font-semibold mb-4 flex items-center gap-2"><Target size={15} className="text-indigo-500" /> Bu Ay Bütçe Takibi <span className="text-xs font-normal text-slate-400">(Ayarlar'dan limit belirleyin)</span></h2>
                  <div className="space-y-3">
                    {rows.map(({ k, harcanan, limit, pct }) => (
                      <div key={k}>
                        <div className="flex items-center justify-between mb-1">
                          <span className="text-xs font-medium">{k}</span>
                          <span className={`text-xs font-medium ${pct >= 100 ? 'text-rose-600 dark:text-rose-400' : pct >= 80 ? 'text-amber-600 dark:text-amber-400' : 'text-slate-500'}`}>{formatTL(harcanan)} / {formatTL(limit)} ({pct}%)</span>
                        </div>
                        <div className="h-2 rounded-full bg-slate-100 dark:bg-slate-700 overflow-hidden">
                          <div className={`h-full rounded-full transition-all ${pct >= 100 ? 'bg-rose-500' : pct >= 80 ? 'bg-amber-500' : 'bg-indigo-500'}`} style={{ width: `${pct}%` }} />
                        </div>
                      </div>
                    ))}
                  </div>
                </Card>
              );
            })()}

            {/* SABİT GİDERLER + OFİS GİDERLERİ */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <SabitGiderlerCard
                sabitGiderler={sabitGiderler}
                onEkle={sg => setSabitGiderler(p => [...p, sg])}
                onSil={id => setSabitGiderler(p => p.filter(sg => sg.id !== id))}
                onUygula={(sg, buAy) => {
                  const kdvTutar = hesaplaKdv(sg.tutar, sg.kdvOrani);
                  setTransactions(p => [{
                    id: Date.now(), tarih: `${buAy}-01`, tur: 'Gider', personel: null,
                    aciklama: sg.ad, kategori: sg.kategori,
                    netTutar: sg.tutar, kdvOrani: sg.kdvOrani, kdvTutar, toplamTutar: sg.tutar + kdvTutar,
                    sabitGiderId: sg.id,
                  }, ...p]);
                }}
                transactions={transactions}
              />
              <OfisGiderleriCard onAdd={r => setTransactions(p => [r, ...p])} />
            </div>

            {/* CHAT + GRAFİKLER */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
              <Card className="flex flex-col h-96">
                <h2 className="text-sm font-semibold mb-3 flex items-center gap-2"><PlusCircle size={15} className="text-indigo-500" /> Asistan Girişi</h2>
                <div className="flex-1 overflow-y-auto space-y-2 mb-3 pr-1">
                  {chatLog.map((msg, i) => (
                    <div key={i} className={`flex ${msg.from === 'user' ? 'justify-end' : 'justify-start'}`}>
                      <div className={`max-w-[90%] rounded-2xl px-3 py-2 text-xs leading-relaxed ${msg.from === 'user' ? 'bg-indigo-600 text-white rounded-br-sm' : 'bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-bl-sm'}`}>
                        {msg.loading ? <span className="animate-pulse">Gemini düşünüyor...</span> : msg.text}
                      </div>
                    </div>
                  ))}
                </div>
                <form onSubmit={handleChatSubmit} className="flex gap-2">
                  <input value={chatInput} onChange={e => setChatInput(e.target.value)} placeholder='"Personel 1, 800 TL ofis harcaması, %20 KDV"' className="flex-1 text-xs border border-slate-200 dark:border-slate-600 rounded-xl px-3 py-2 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 placeholder-slate-400" />
                  <button type="submit" className="p-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl"><Send size={14} /></button>
                </form>
              </Card>
              <Card>
                <h2 className="text-sm font-semibold mb-3">Personel Harcamaları</h2>
                <ResponsiveContainer width="100%" height={260}>
                  <PieChart>
                    <Pie data={personelHarcama} dataKey="toplam" nameKey="name" cx="50%" cy="50%" innerRadius={60} outerRadius={90} paddingAngle={3}>
                      {personelHarcama.map((_, i) => <Cell key={i} fill={RENK_PALETI[i % RENK_PALETI.length]} />)}
                    </Pie>
                    <Tooltip formatter={v => formatTL(v)} />
                    <Legend iconType="circle" iconSize={8} />
                  </PieChart>
                </ResponsiveContainer>
              </Card>
              <Card>
                <h2 className="text-sm font-semibold mb-3">Kategori Dağılımı (Gider)</h2>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={kategoriDagilim} layout="vertical" margin={{ left: 20, right: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={isDark ? '#334155' : '#e2e8f0'} />
                    <XAxis type="number" tick={{ fontSize: 10 }} tickFormatter={v => `${(v / 1000).toFixed(0)}K`} />
                    <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={70} />
                    <Tooltip formatter={v => formatTL(v)} />
                    <Bar dataKey="value" radius={[0, 6, 6, 0]}>{kategoriDagilim.map((_, i) => <Cell key={i} fill={RENK_PALETI[i % RENK_PALETI.length]} />)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>

            {/* AYLIK TREND (tüm veri — tarih filtresi uygulanmaz) */}
            <Card>
              <h2 className="text-sm font-semibold mb-4 flex items-center gap-2"><TrendingUp size={15} className="text-indigo-500" /> Aylık Gelir-Gider Trendi <span className="text-xs font-normal text-slate-400">(tüm dönem)</span></h2>
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={aylikTrend} margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={isDark ? '#334155' : '#e2e8f0'} />
                  <XAxis dataKey="ay" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${(v / 1000).toFixed(0)}K`} width={45} />
                  <Tooltip formatter={(v, n) => [formatTL(v), n === 'gelir' ? 'Gelir' : n === 'gider' ? 'Gider' : 'Net Kar']} contentStyle={{ fontSize: 11, borderRadius: 8, border: isDark ? '1px solid #334155' : '1px solid #e2e8f0', background: isDark ? '#1e293b' : '#fff', color: isDark ? '#f1f5f9' : '#0f172a' }} />
                  <Legend iconType="circle" iconSize={8} formatter={v => v === 'gelir' ? 'Gelir' : v === 'gider' ? 'Gider' : 'Net Kar'} />
                  <Line type="monotone" dataKey="gelir" stroke="#10b981" strokeWidth={2.5} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                  <Line type="monotone" dataKey="gider" stroke="#f43f5e" strokeWidth={2.5} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                  <Line type="monotone" dataKey="kar"   stroke="#6366f1" strokeWidth={2} strokeDasharray="5 3" dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </Card>

            {/* TABLO */}
            <Card className="overflow-hidden p-0">
              <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex flex-wrap gap-3 items-center justify-between">
                <h2 className="text-sm font-semibold flex items-center gap-2"><Filter size={14} className="text-slate-400" /> Gelir-Gider Tablosu <Badge color="violet">{filteredRows.length} kayıt</Badge></h2>
                <div className="flex flex-wrap gap-2">
                  {[
                    { label: 'Tür',      value: filterTur,      setter: setFilterTur,      options: ['Tümü', 'Gelir', 'Gider'] },
                    { label: 'Personel', value: filterPersonel, setter: setFilterPersonel, options: ['Tümü', ...personeller] },
                    { label: 'Kategori', value: filterKategori, setter: setFilterKategori, options: ['Tümü', ...KATEGORILER] },
                  ].map(f => (
                    <select key={f.label} value={f.value} onChange={e => f.setter(e.target.value)} className="text-xs border border-slate-200 dark:border-slate-600 rounded-lg px-2 py-1.5 bg-white dark:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500">
                      {f.options.map(o => <option key={o}>{o}</option>)}
                    </select>
                  ))}
                  <button onClick={() => { setFilterTur('Tümü'); setFilterPersonel('Tümü'); setFilterKategori('Tümü'); }} className="text-xs flex items-center gap-1 px-2 py-1.5 rounded-lg border border-slate-200 dark:border-slate-600 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"><Trash2 size={11} /> Sıfırla</button>
                  <button onClick={() => exportMuhasebeExcel(filteredRows, datePreset !== 'tumu' ? datePreset : '')} className="text-xs flex items-center gap-1.5 px-2 py-1.5 rounded-lg border border-emerald-300 dark:border-emerald-700 text-emerald-700 dark:text-emerald-400 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 transition-colors font-medium"><FileSpreadsheet size={11} /> Rapor İndir</button>
                </div>
              </div>
              {(() => {
                const hasFatura = filteredRows.some(r => r.faturaNo);
                const cols = ['Tarih', 'Tür', 'Personel', 'Açıklama', ...(hasFatura ? ['Fatura No'] : []), 'Kategori', 'Net Tutar', 'KDV %', 'KDV Tutarı', 'Toplam', ''];
                return (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-slate-500 dark:text-slate-400">{cols.map(h => <th key={h} className="px-4 py-3 text-left font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
                      <tbody>
                        {filteredRows.map((row, i) => (
                          <tr key={row.id} className={`border-t border-slate-100 dark:border-slate-700/60 hover:bg-slate-50 dark:hover:bg-slate-800/40 transition-colors ${i % 2 ? 'bg-slate-50/50 dark:bg-slate-800/20' : ''}`}>
                            <td className="px-4 py-3 text-slate-500 whitespace-nowrap">{row.tarih}</td>
                            <td className="px-4 py-3"><Badge color={row.tur === 'Gelir' ? 'green' : 'red'}>{row.tur}</Badge></td>
                            <td className="px-4 py-3 text-slate-500">{row.personel || '—'}</td>
                            <td className="px-4 py-3 max-w-[200px] truncate" title={row.aciklama}>{row.aciklama}</td>
                            {hasFatura && <td className="px-4 py-3 text-slate-400 font-mono text-[11px]">{row.faturaNo || '—'}</td>}
                            <td className="px-4 py-3"><Badge color="blue">{row.kategori}</Badge></td>
                            <td className="px-4 py-3 font-medium">{formatTL(row.netTutar)}</td>
                            <td className="px-4 py-3"><Badge color="amber">%{row.kdvOrani}</Badge></td>
                            <td className="px-4 py-3 text-slate-500">{formatTL(row.kdvTutar)}</td>
                            <td className="px-4 py-3 font-semibold">{formatTL(row.toplamTutar)}</td>
                            <td className="px-4 py-3"><button onClick={() => setTransactions(p => p.filter(t => t.id !== row.id))} className="p-1 rounded text-slate-300 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"><X size={12} /></button></td>
                          </tr>
                        ))}
                        {filteredRows.length === 0 && <tr><td colSpan={cols.length} className="px-4 py-8 text-center text-slate-400">Kayıt bulunamadı</td></tr>}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </Card>
          </>)}

          {/* ── KASA ── */}
          {activeTab === 'kasa' && (
            <KasaTab
              kasaHareketleri={kasaHareketleri}
              aylikTrend={aylikTrend}
              onKasaEkle={h => setKasaHareketleri(p => [h, ...p])}
              onKasaDelete={id => setKasaHareketleri(p => p.filter(h => h.id !== id))}
            />
          )}

          {/* ── KDV BEYANNAME ── */}
          {activeTab === 'kdv' && <KdvTab transactions={transactions} />}

          {/* ── PROFORMA FATURA ── */}
          {activeTab === 'proforma' && (
            <ProformaTab
              proformalar={proformalar}
              musteriler={musteriler}
              logo={logo}
              onAdd={p => setProformalar(prev => [p, ...prev])}
              onUpdate={p => setProformalar(prev => prev.map(pr => pr.id === p.id ? p : pr))}
              onDelete={id => setProformalar(prev => prev.filter(p => p.id !== id))}
              onAktar={p => {
                const byKdv = {};
                p.kalemler.forEach(k => { const kh = calcKalem(k); if (!byKdv[k.kdvOrani]) byKdv[k.kdvOrani] = { net: 0, kdv: 0 }; byKdv[k.kdvOrani].net += kh.netTutar; byKdv[k.kdvOrani].kdv += kh.kdvTutar; });
                const newTx = Object.entries(byKdv).map(([oran, vals], idx) => ({ id: Date.now() + idx, tarih: p.tarih, tur: 'Gelir', personel: null, aciklama: `${p.faturaNo} — ${p.musteri.ad}`, kategori: 'Diğer', netTutar: vals.net, kdvOrani: parseInt(oran), kdvTutar: vals.kdv, toplamTutar: vals.net + vals.kdv, faturaNo: p.faturaNo }));
                setTransactions(prev => [...newTx, ...prev]);
                setProformalar(prev => prev.map(pr => pr.id === p.id ? { ...pr, aktarildi: true, durum: 'Onaylandı' } : pr));
              }}
            />
          )}

          {/* ── AYARLAR ── */}
          {activeTab === 'ayarlar' && (
            <AyarlarTab
              personeller={personeller}
              onPersonelEkle={isim => setPersoneller(p => [...p, isim])}
              onPersonelSil={isim => setPersoneller(p => p.filter(x => x !== isim))}
              transactions={transactions}
              kasaHareketleri={kasaHareketleri}
              onResetAll={handleResetAll}
              butceler={butceler}
              onButceChange={(kat, val) => setButceler(b => ({ ...b, [kat]: val }))}
              musteriler={musteriler}
              onMusteriEkle={m => setMusteriler(p => [...p, m])}
              onMusteriSil={id => setMusteriler(p => p.filter(m => m.id !== id))}
              logo={logo}
              onLogoChange={setLogo}
            />
          )}

        </main>

        <footer className="text-center text-xs text-slate-400 py-4">
          3065 Sayılı KDV Kanunu uyarınca: Ödenecek KDV = Alınan KDV − Ödenen KDV
        </footer>
      </div>
    </div>
  );
}
