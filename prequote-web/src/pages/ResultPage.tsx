import { useEffect, useState, useCallback, useRef, type ReactNode } from 'react';
import { useSearchParams } from 'react-router-dom';
import { getResultData, getResultRecommendations } from '../api/gasClient';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import {
  Phone,
  MessageCircle,
  CalendarCheck,
  AlertTriangle,
  CheckCircle2,
  Info,
  ArrowRight,
  Sparkles,
  ClipboardList,
  ShieldCheck,
  Package,
  ChevronRight,
} from 'lucide-react';

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

interface RequestInfo {
  customer_name?: string;
  created_at?: string;
  project_type_label?: string;
  housing_type_label?: string;
  area_label?: string;
  scope_level_label?: string;
  flow_mode_label?: string;
  schedule_start_label?: string;
  schedule_end_label?: string;
  address_label?: string;
  customer_note_label?: string;
}

interface Estimate {
  min?: number;
  max?: number;
}

interface Flag {
  severity?: string;
  severity_label?: string;
  label?: string;
  flag?: string;
}

interface AnswerItem {
  question_label?: string;
  question_code?: string;
  answer_display_text?: string;
  answer_display_values?: string[];
}

interface AnswerGroup {
  group_label?: string;
  group_key?: string;
  items?: AnswerItem[];
}

interface Recommendation {
  rec_id?: string;
  title?: string;
  material_id?: string;
  brand?: string;
  spec?: string;
  subtitle?: string;
  image_url?: string;
  reason_text?: string;
  price_hint_min?: number;
  price_hint_max?: number;
  tags?: string[];
  trade_code?: string;
  material_type?: string;
}

interface ResultData {
  request?: RequestInfo;
  estimate?: Estimate;
  flags?: Flag[];
  answer_groups?: AnswerGroup[];
  recommendations?: Recommendation[];
  recommendation_status?: string;
  notice_text?: string;
  _bootstrap?: boolean;
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

const BOOTSTRAP_PREFIX = 'pq_result_bootstrap:';
const CACHE_PREFIX = 'pq_result_data:';

function num(v: unknown): number {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function formatMan(v: unknown): string {
  return Math.round(num(v) / 10000).toLocaleString('ko-KR');
}

function formatRange(minV: unknown, maxV: unknown): string {
  return `${formatMan(minV)} ~ ${formatMan(maxV)}만원`;
}

function formatDateTime(v: unknown): string {
  if (!v) return '-';
  const d = new Date(v as string);
  if (isNaN(d.getTime())) return String(v);
  return new Intl.DateTimeFormat('ko-KR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(d);
}

function readJson(storage: Storage, key: string) {
  try {
    const raw = storage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeJson(storage: Storage, key: string, value: unknown) {
  try {
    storage.setItem(key, JSON.stringify(value ?? null));
  } catch { /* quota etc */ }
}

function rangePercent(min: number, max: number): { left: number; width: number } {
  const lo = Math.round(min / 10000);
  const hi = Math.round(max / 10000);
  if (hi <= 0) return { left: 20, width: 60 };
  const pad = Math.max(hi * 0.15, 200);
  const totalRange = hi + pad;
  return {
    left: Math.max((lo / totalRange) * 100, 5),
    width: Math.min(((hi - lo) / totalRange) * 100, 80),
  };
}

/* ------------------------------------------------------------------ */
/*  Severity helpers                                                   */
/* ------------------------------------------------------------------ */

function severityColor(severity?: string) {
  const s = (severity || '').toUpperCase();
  if (s === 'HIGH') return { bg: 'bg-red-50', border: 'border-red-200', dot: 'bg-red-500', text: 'text-red-700' };
  if (s === 'MEDIUM') return { bg: 'bg-amber-50', border: 'border-amber-200', dot: 'bg-amber-500', text: 'text-amber-700' };
  return { bg: 'bg-emerald-50', border: 'border-emerald-200', dot: 'bg-emerald-500', text: 'text-emerald-700' };
}

/* ------------------------------------------------------------------ */
/*  Component                                                          */
/* ------------------------------------------------------------------ */

export default function ResultPage() {
  const [searchParams] = useSearchParams();
  const id = searchParams.get('id') || '';
  const token = searchParams.get('token') || '';

  const [data, setData] = useState<ResultData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const hasFetched = useRef(false);

  // consume bootstrap + cached data
  const consumeBootstrap = useCallback(() => {
    if (!id || !token) return null;
    const bsKey = `${BOOTSTRAP_PREFIX}${id}:${token}`;
    const cached = readJson(sessionStorage, bsKey);
    if (cached) {
      try { sessionStorage.removeItem(bsKey); } catch { /* noop */ }
      return cached as ResultData;
    }
    const cacheKey = `${CACHE_PREFIX}${id}:${token}`;
    return readJson(localStorage, cacheKey) as ResultData | null;
  }, [id, token]);

  // persist to cache
  const persist = useCallback(
    (d: ResultData) => {
      if (!id || !token || d._bootstrap) return;
      writeJson(localStorage, `${CACHE_PREFIX}${id}:${token}`, d);
    },
    [id, token],
  );

  // fetch main data
  const fetchData = useCallback(async () => {
    try {
      const result = (await getResultData(id, token)) as ResultData;
      setData(result);
      persist(result);
      setLoading(false);
      return result;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : '결과를 불러오지 못했습니다.';
      setError(msg);
      setLoading(false);
      return null;
    }
  }, [id, token, persist]);

  // lazy-load recommendations
  const hydrateRecs = useCallback(
    async (current: ResultData) => {
      try {
        const res = (await getResultRecommendations(id, token)) as {
          recommendations?: Recommendation[];
        };
        const updated: ResultData = {
          ...current,
          recommendations: res?.recommendations || [],
          recommendation_status: 'READY',
        };
        setData(updated);
        persist(updated);
      } catch { /* silent */ }
    },
    [id, token, persist],
  );

  useEffect(() => {
    if (hasFetched.current) return;
    hasFetched.current = true;

    if (!id || !token) {
      setError('유효한 결과 링크가 아닙니다.');
      setLoading(false);
      return;
    }

    const initial = consumeBootstrap();
    if (initial) {
      setData(initial);
      setLoading(false);
      // background refresh
      fetchData().then((fresh) => {
        if (
          fresh &&
          !fresh._bootstrap &&
          (fresh.recommendation_status || '').toUpperCase() === 'PENDING'
        ) {
          setTimeout(() => hydrateRecs(fresh), 200);
        }
      });
    } else {
      fetchData().then((fresh) => {
        if (
          fresh &&
          !fresh._bootstrap &&
          (fresh.recommendation_status || '').toUpperCase() === 'PENDING'
        ) {
          setTimeout(() => hydrateRecs(fresh), 200);
        }
      });
    }
  }, [id, token, consumeBootstrap, fetchData, hydrateRecs]);

  /* ---------- Loading state ---------- */
  if (loading) {
    return (
      <div className="min-h-screen bg-[#f5f3f0]">
        <div className="mx-auto max-w-[940px] px-5 py-10">
          <div className="animate-fade-in space-y-5">
            <Skeleton className="h-[260px] w-full rounded-[28px]" />
            <div className="grid gap-5 md:grid-cols-2">
              <Skeleton className="h-[200px] rounded-[28px]" />
              <Skeleton className="h-[200px] rounded-[28px]" />
            </div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- Error state ---------- */
  if (error) {
    return (
      <div className="min-h-screen bg-[#f5f3f0] flex items-center justify-center px-5">
        <div className="w-full max-w-md bg-white/95 rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-lg p-8 text-center animate-fade-in">
          <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-full bg-red-50">
            <AlertTriangle className="h-7 w-7 text-red-500" />
          </div>
          <h2 className="font-serif text-xl mb-2">결과를 불러오지 못했습니다</h2>
          <p className="text-sm text-[#696969]">{error}</p>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const req = data.request || {};
  const est = data.estimate || {};
  const flags = data.flags || [];
  const answers = data.answer_groups || [];
  const recs = data.recommendations || [];
  const recStatus = (data.recommendation_status || '').toUpperCase();
  const notice =
    data.notice_text ||
    '입력하신 조건을 기준으로 참고 범위를 먼저 보여드립니다. 최종 금액은 상담 후 다시 안내됩니다.';

  const rangeBar = rangePercent(num(est.min), num(est.max));
  const scopeBadges = [
    req.project_type_label,
    req.housing_type_label,
    req.area_label,
    req.scope_level_label,
  ].filter(Boolean);

  return (
    <div className="min-h-screen bg-[#f5f3f0]">
      <div className="mx-auto max-w-[940px] px-5 py-7 pb-24 md:py-10">
        <div className="space-y-5 animate-fade-in">
          {/* ============================================================ */}
          {/*  HERO BAND                                                   */}
          {/* ============================================================ */}
          <section className="grid gap-5 md:grid-cols-[1.08fr_0.92fr]">
            {/* Left: Copy */}
            <div className="bg-white/[0.94] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6 md:p-8">
              <span className="inline-flex items-center gap-1.5 rounded-full bg-[#8b6d4b]/10 px-3 py-1 text-xs font-bold text-[#8b6d4b]">
                <Sparkles className="h-3.5 w-3.5" />
                가견적 결과
              </span>
              <h1 className="mt-4 mb-3 font-serif text-[clamp(1.8rem,4vw,2.6rem)] leading-[1.05] tracking-tight text-[#2c2c2c]">
                먼저 금액 범위를 확인하세요
              </h1>
              <p className="text-sm leading-relaxed text-[#696969]">
                결과 요약을 먼저 보여드리고, 상세 응답과 추천 자재는 아래에서 이어서 정리합니다.
              </p>
              <div className="mt-5 rounded-2xl border border-[#8b6d4b]/15 bg-[#8b6d4b]/5 px-4 py-3 text-sm leading-relaxed text-[#696969]">
                {notice}
              </div>
            </div>

            {/* Right: Range */}
            <div className="bg-gradient-to-br from-[#fffdf8] to-[#f3ead8] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6 md:p-8 flex flex-col justify-between">
              <span className="inline-flex items-center gap-1.5 rounded-full bg-[#8b6d4b]/10 px-3 py-1 text-xs font-bold text-[#8b6d4b]">
                현재 참고 범위
              </span>
              <div className="mt-4">
                <div className="font-serif text-[clamp(2rem,5vw,3.2rem)] font-bold leading-none text-[#2c2c2c] tracking-tight">
                  {formatRange(est.min || 0, est.max || 0)}
                </div>

                {/* Range bar visualisation */}
                <div className="mt-5 mb-2">
                  <div className="relative h-3 w-full rounded-full bg-[#8b6d4b]/10 overflow-hidden">
                    <div
                      className="absolute inset-y-0 rounded-full bg-gradient-to-r from-[#8b6d4b] to-[#b9a79e] transition-all duration-700"
                      style={{ left: `${rangeBar.left}%`, width: `${rangeBar.width}%` }}
                    />
                  </div>
                  <div className="mt-1 flex justify-between text-[10px] text-[#696969]">
                    <span>{formatMan(est.min || 0)}만원</span>
                    <span>{formatMan(est.max || 0)}만원</span>
                  </div>
                </div>
              </div>

              <div className="mt-4 space-y-2.5">
                <div className="flex justify-between text-sm">
                  <span className="text-[#696969]">고객명</span>
                  <span className="font-medium text-[#333]">{req.customer_name || '고객'}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-[#696969]">접수 시각</span>
                  <span className="font-medium text-[#333]">{formatDateTime(req.created_at)}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-[#696969]">프로젝트</span>
                  <span className="font-medium text-[#333]">
                    {[req.project_type_label, req.housing_type_label].filter(Boolean).join(' / ') || '-'}
                  </span>
                </div>
              </div>
            </div>
          </section>

          {/* ============================================================ */}
          {/*  4-STEP FLOW GRID                                            */}
          {/* ============================================================ */}
          <div className="grid gap-5 md:grid-cols-2">
            {/* ---- STEP 1: 접수 요약 ---- */}
            <section className="bg-white/[0.94] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6">
              <SectionHead step={1} icon={<ClipboardList className="h-4 w-4" />} label="STEP 1" title="접수 요약" desc="입력하신 기본 정보를 먼저 정리했습니다." />
              <div className="space-y-2.5">
                {scopeBadges.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 mb-3">
                    {scopeBadges.map((b, i) => (
                      <Badge key={i} variant="secondary" className="text-[11px] bg-[#8b6d4b]/8 text-[#8b6d4b] border-[#8b6d4b]/15">
                        {b}
                      </Badge>
                    ))}
                  </div>
                )}
                <InfoRow label="프로젝트" value={[req.project_type_label, req.housing_type_label].filter(Boolean).join(' / ')} />
                <InfoRow label="면적/범위" value={[req.area_label, req.scope_level_label].filter(Boolean).join(' / ')} />
                <InfoRow label="진행 방식" value={req.flow_mode_label} />
                <InfoRow
                  label="희망 일정"
                  value={`${req.schedule_start_label || '미정'} ~ ${req.schedule_end_label || '미정'}`}
                />
                <InfoRow label="주소" value={req.address_label} />
                <InfoRow label="요청 메모" value={req.customer_note_label || '없음'} />
              </div>
            </section>

            {/* ---- STEP 2: 체크 포인트 ---- */}
            <section className="bg-white/[0.94] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6">
              <SectionHead step={2} icon={<ShieldCheck className="h-4 w-4" />} label="STEP 2" title="체크 포인트" desc="추가 확인이 필요한 내용입니다." />
              {flags.length === 0 ? (
                <EmptyBox text="추가 확인이 필요한 항목이 없습니다." />
              ) : (
                <div className="space-y-2.5">
                  {flags.map((flag, i) => {
                    const c = severityColor(flag.severity);
                    return (
                      <div
                        key={i}
                        className={`flex items-start gap-3 rounded-2xl border ${c.border} ${c.bg} px-4 py-3 transition-all duration-300`}
                      >
                        <span className={`mt-1 h-2.5 w-2.5 shrink-0 rounded-full ${c.dot}`} />
                        <div>
                          <span className={`text-xs font-bold ${c.text}`}>
                            {flag.severity_label || flag.severity || '참고'}
                          </span>
                          <p className="mt-1 text-sm text-[#696969] leading-relaxed">
                            {flag.label || flag.flag}
                          </p>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </section>

            {/* ---- STEP 3: 추천 자재 ---- */}
            <section className="bg-white/[0.94] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6">
              <SectionHead step={3} icon={<Package className="h-4 w-4" />} label="STEP 3" title="추천 자재" desc="저장된 추천을 먼저 보여줍니다." />
              {recStatus === 'PENDING' ? (
                <div className="space-y-3">
                  <Skeleton className="h-20 rounded-2xl" />
                  <Skeleton className="h-20 rounded-2xl" />
                  <p className="text-center text-xs text-[#696969] mt-2">추천 자재를 준비하는 중입니다...</p>
                </div>
              ) : recs.length === 0 ? (
                <EmptyBox text="아직 추천 자재가 준비되지 않았습니다. 상담 시 맞춤 자재를 안내해드립니다." />
              ) : (
                <div className="space-y-3">
                  {recs.map((item, i) => (
                    <RecommendationCard key={item.rec_id || i} item={item} />
                  ))}
                </div>
              )}
            </section>

            {/* ---- STEP 4: 설문 응답 ---- */}
            <section className="bg-white/[0.94] rounded-[28px] border border-[rgba(122,116,108,.14)] shadow-[0_18px_42px_rgba(32,29,24,.08)] p-6">
              <SectionHead step={4} icon={<Info className="h-4 w-4" />} label="STEP 4" title="설문 응답" desc="입력한 내용을 다시 확인할 수 있습니다." />
              {answers.length === 0 ? (
                <EmptyBox text="설문 응답을 정리하는 중입니다." />
              ) : (
                <div className="space-y-4">
                  {answers.map((group, gi) => (
                    <div key={gi}>
                      <h4 className="font-serif text-base font-medium text-[#333] mb-2">
                        {group.group_label || group.group_key || '응답'}
                      </h4>
                      <div className="space-y-2">
                        {(group.items || []).map((item, ai) => (
                          <div
                            key={ai}
                            className="rounded-2xl border border-[rgba(122,116,108,.12)] bg-white px-4 py-3"
                          >
                            <span className="block text-xs font-bold text-[#696969] mb-1">
                              {item.question_label || item.question_code || '-'}
                            </span>
                            {(item.answer_display_values || []).length > 1 ? (
                              <div className="flex flex-wrap gap-1.5">
                                {item.answer_display_values!.map((v, vi) => (
                                  <Badge key={vi} variant="secondary" className="text-[11px]">
                                    {v}
                                  </Badge>
                                ))}
                              </div>
                            ) : (
                              <p className="text-sm text-[#333]">
                                {item.answer_display_text || '미입력'}
                              </p>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </section>
          </div>

          {/* ============================================================ */}
          {/*  TRUST MESSAGE                                               */}
          {/* ============================================================ */}
          <div className="flex items-center justify-center gap-2 text-sm text-[#696969] py-2">
            <CheckCircle2 className="h-4 w-4 text-[#8b6d4b]" />
            이 범위는 실제 상담 시 확정됩니다
          </div>

          {/* ============================================================ */}
          {/*  CTA                                                         */}
          {/* ============================================================ */}
          <section className="bg-gradient-to-br from-[#8b6d4b] to-[#6b5139] rounded-[28px] shadow-[0_18px_42px_rgba(32,29,24,.16)] p-8 md:p-10 text-center text-white">
            <h2 className="font-serif text-[clamp(1.4rem,3vw,2rem)] leading-tight mb-3">
              무료 상담을 예약하세요
            </h2>
            <p className="text-white/70 text-sm mb-6 max-w-md mx-auto">
              전문 인테리어 컨설턴트가 맞춤 견적과 자재를 안내해 드립니다.
            </p>
            <div className="flex flex-col sm:flex-row items-center justify-center gap-3">
              <Button
                size="lg"
                className="bg-white text-[#8b6d4b] hover:bg-white/90 font-semibold rounded-full px-8 h-12 shadow-lg transition-all duration-300 hover:shadow-xl hover:-translate-y-0.5"
              >
                <CalendarCheck className="h-4 w-4 mr-2" />
                무료 상담 예약하기
                <ArrowRight className="h-4 w-4 ml-1" />
              </Button>
              <div className="flex items-center gap-4 text-sm text-white/70">
                <a href="tel:02-1234-5678" className="flex items-center gap-1.5 hover:text-white transition-colors">
                  <Phone className="h-3.5 w-3.5" />
                  전화 상담
                </a>
                <a href="#" className="flex items-center gap-1.5 hover:text-white transition-colors">
                  <MessageCircle className="h-3.5 w-3.5" />
                  카카오톡
                </a>
              </div>
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Sub-components                                                     */
/* ------------------------------------------------------------------ */

function SectionHead({
  step,
  icon,
  label,
  title,
  desc,
}: {
  step: number;
  icon: ReactNode;
  label: string;
  title: string;
  desc: string;
}) {
  return (
    <div className="flex items-start justify-between gap-3 mb-5">
      <div>
        <span className="inline-flex items-center gap-1.5 text-[11px] font-bold tracking-widest text-[#8b6d4b]">
          {icon}
          {label}
        </span>
        <h3 className="font-serif text-lg mt-1 text-[#2c2c2c]">{title}</h3>
        <p className="text-xs text-[#696969] mt-1 max-w-[280px]">{desc}</p>
      </div>
      <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-[#8b6d4b]/8 text-[#8b6d4b] text-sm font-bold font-serif">
        {step}
      </span>
    </div>
  );
}

function InfoRow({ label, value }: { label: string; value?: string }) {
  return (
    <div className="grid grid-cols-[90px_1fr] gap-3 text-sm">
      <span className="text-xs font-bold text-[#696969] tracking-wide pt-0.5">{label}</span>
      <span className="text-[#333]">{value || '미입력'}</span>
    </div>
  );
}

function EmptyBox({ text }: { text: string }) {
  return (
    <div className="rounded-2xl border border-dashed border-[rgba(122,116,108,.18)] bg-white/60 px-5 py-6 text-center text-sm text-[#696969]">
      {text}
    </div>
  );
}

function RecommendationCard({ item }: { item: Recommendation }) {
  const summary = [item.brand, item.spec, item.subtitle].filter(Boolean).join(' / ');
  const tags = [item.trade_code, item.material_type, ...(item.tags || [])].filter(Boolean);

  return (
    <div className="group flex gap-3 rounded-2xl border border-[rgba(122,116,108,.12)] bg-white px-4 py-3 transition-all duration-200 hover:shadow-md hover:-translate-y-0.5">
      {/* Thumbnail */}
      <div
        className="h-[72px] w-[72px] shrink-0 rounded-2xl bg-gradient-to-br from-[#ebe6dc] to-[#d8d1c4] bg-cover bg-center"
        style={item.image_url ? { backgroundImage: `url('${item.image_url}')` } : undefined}
      />
      {/* Content */}
      <div className="min-w-0 flex-1">
        <strong className="block text-sm font-bold font-serif text-[#2c2c2c] leading-tight">
          {item.title || item.material_id || '추천 자재'}
        </strong>
        <p className="mt-1 text-xs text-[#696969] leading-relaxed truncate">
          {summary || '정보 준비 중'}
        </p>
        {item.price_hint_min || item.price_hint_max ? (
          <p className="mt-1 text-[11px] font-medium text-[#8b6d4b]">
            {formatRange(item.price_hint_min || 0, item.price_hint_max || 0)}
          </p>
        ) : null}
        {tags.length > 0 && (
          <div className="mt-2 flex flex-wrap gap-1">
            {tags.map((t, i) => (
              <span
                key={i}
                className="inline-block rounded-full bg-[rgba(122,116,108,.08)] px-2 py-0.5 text-[10px] text-[#696969]"
              >
                {t}
              </span>
            ))}
          </div>
        )}
        {item.reason_text && (
          <p className="mt-1.5 text-xs text-[#696969] leading-relaxed">{item.reason_text}</p>
        )}
      </div>
      <ChevronRight className="mt-5 h-4 w-4 shrink-0 text-[#b9a79e] opacity-0 group-hover:opacity-100 transition-opacity" />
    </div>
  );
}
