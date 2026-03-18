import { useEffect, useRef, useState } from 'react';
import type { SurveyConfig } from '@/survey/types';

interface SurveyLandingProps {
  config: SurveyConfig | null;
  loading: boolean;
  onStart: () => void;
}

/* ── Fallback content (hardcoded Korean) ─────────────────────────────────── */

const FALLBACK_TAG = '전문 인테리어 컨설팅';
const FALLBACK_TITLE = '3분 안에 받아보는\n맞춤 인테리어 견적';
const FALLBACK_DESC =
  '10년 이상 경력의 전문 컨설턴트가 직접 검토하는 맞춤 견적입니다.\n간단한 정보만 입력하시면 예상 비용과 추천 자재를 바로 확인하실 수 있습니다.';
const FALLBACK_CTA = '무료 맞춤 견적 받기';

const TRUST_BADGES = ['평균 3분 소요', '전문가 직접 검토', '무료 상담 포함'];

const POINTS = [
  {
    title: '정확한 범위 견적',
    body: '실측 전이라도 평형과 공사 범위만으로 신뢰할 수 있는 견적 범위를 제공합니다.',
  },
  {
    title: '전문가 맞춤 검토',
    body: '제출된 내용은 인테리어 전문 컨설턴트가 직접 확인하고 맞춤 안내드립니다.',
  },
  {
    title: '필요한 것만 간결하게',
    body: '선택하지 않은 공정은 자동으로 건너뛰어 불필요한 질문이 없습니다.',
  },
];

const TRUST_METRICS = [
  { end: 3200, suffix: '+', label: '누적 상담 건수' },
  { end: 98, suffix: '%', label: '고객 만족도' },
  { end: 10, suffix: '년+', label: '전문 경력' },
];

/* ── UI helpers from config ──────────────────────────────────────────────── */

function uiHero(config: SurveyConfig | null) {
  const list = config?.uiContent || [];
  return list.find((u) => u.content_key === 'landing_hero') || {};
}

function uiChips(config: SurveyConfig | null) {
  return (config?.uiContent || [])
    .filter((u) => u.content_type === 'CHIP')
    .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
}

function uiPoints(config: SurveyConfig | null) {
  return (config?.uiContent || [])
    .filter((u) => u.content_type === 'POINT')
    .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
}

/* ── Count-up hook ───────────────────────────────────────────────────────── */

function useCountUp(end: number, duration = 1800): number {
  const [value, setValue] = useState(0);
  const rafRef = useRef<number>(0);

  useEffect(() => {
    const start = performance.now();
    const tick = (now: number) => {
      const elapsed = now - start;
      const progress = Math.min(elapsed / duration, 1);
      // Ease-out cubic
      const eased = 1 - Math.pow(1 - progress, 3);
      setValue(Math.round(eased * end));
      if (progress < 1) {
        rafRef.current = requestAnimationFrame(tick);
      }
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [end, duration]);

  return value;
}

/* ── Metric counter component ────────────────────────────────────────────── */

function MetricCounter({
  end,
  suffix,
  label,
}: {
  end: number;
  suffix: string;
  label: string;
}) {
  const count = useCountUp(end);
  return (
    <div className="text-center">
      <div className="font-sans text-2xl font-bold leading-tight text-[#8b6d4b]">
        {count}
        <span className="text-lg">{suffix}</span>
      </div>
      <div className="font-sans text-[11px] text-[#999] mt-1">{label}</div>
    </div>
  );
}

/* ── Main component ──────────────────────────────────────────────────────── */

export default function SurveyLanding({
  config,
  loading,
  onStart,
}: SurveyLandingProps) {
  const hero = uiHero(config);
  const chips = uiChips(config);
  const points = uiPoints(config);

  // Use config values with hardcoded fallbacks
  const tag = FALLBACK_TAG;
  const title =
    hero.title || config?.settings?.hero_title || FALLBACK_TITLE;
  const desc =
    hero.body || config?.settings?.hero_subtitle || FALLBACK_DESC;
  const ctaLabel =
    (hero as Record<string, string>).cta_label || FALLBACK_CTA;

  const displayChips =
    chips.length > 0 ? chips.map((c) => c.title) : TRUST_BADGES;
  const displayPoints = points.length > 0 ? points : POINTS;

  return (
    <div className="min-h-screen flex flex-col animate-fade-in bg-[#faf9f7]">
      {/* Hero area with subtle background image */}
      <div className="relative flex-1 flex flex-col items-center justify-center px-5 pt-12 pb-6 text-center overflow-hidden">
        {/* Background hero image at very low opacity */}
        <div
          className="absolute inset-0 bg-cover bg-center pointer-events-none"
          style={{
            backgroundImage: 'url(/images/hero-interior.jpg)',
            opacity: 0.04,
          }}
        />

        {/* Content over image */}
        <div className="relative z-10 w-full max-w-[440px] mx-auto">
          {/* Catchphrase */}
          <p className="font-serif text-sm text-[#b9a79e] tracking-[0.15em] mb-3 italic">
            이상이 일상이 되는 시간
          </p>

          {/* Tag */}
          <span className="inline-block font-sans text-xs font-semibold tracking-[0.18em] text-[#8b6d4b] uppercase mb-5 px-3 py-1.5 rounded-full bg-[#8b6d4b]/[0.08]">
            {tag}
          </span>

          {/* Title */}
          <h1 className="font-serif text-[28px] leading-[1.35] tracking-tight mb-4 text-[#2a2a2a] whitespace-pre-line">
            {title}
          </h1>

          {/* Description */}
          <p className="font-sans text-sm font-light text-[#696969] mx-auto mb-7 leading-[1.8] whitespace-pre-line max-w-[380px]">
            {desc}
          </p>

          {/* Trust badges (pills) */}
          <div className="flex gap-2 justify-center flex-wrap mb-8">
            {displayChips.map((text, i) => (
              <span
                key={i}
                className="inline-flex items-center gap-1.5 rounded-full bg-white px-3.5 py-1.5 text-xs font-sans font-medium text-[#696969] shadow-sm border border-[#f0ece8]"
              >
                <span className="w-1.5 h-1.5 rounded-full bg-[#8b6d4b]/40" />
                {text}
              </span>
            ))}
          </div>

          {/* Trust metrics with count-up */}
          <div className="flex gap-8 justify-center mb-8">
            {TRUST_METRICS.map((m, i) => (
              <MetricCounter
                key={i}
                end={m.end}
                suffix={m.suffix}
                label={m.label}
              />
            ))}
          </div>

          {/* Divider */}
          <div className="w-10 h-0.5 bg-[#8b6d4b] opacity-30 mx-auto mb-8" />

          {/* Points card */}
          <div className="bg-white rounded-2xl shadow-sm border border-[#f0ece8] p-6 mb-8">
            <div className="grid gap-5 text-left">
              {displayPoints.map((p, i) => (
                <div key={i} className="flex gap-3.5 items-start">
                  <div className="w-8 h-8 rounded-xl bg-gradient-to-br from-[#8b6d4b]/15 to-[#8b6d4b]/5 text-[#8b6d4b] flex items-center justify-center shrink-0 font-sans font-bold text-sm">
                    {i + 1}
                  </div>
                  <div>
                    <strong className="block font-sans text-[13px] font-semibold text-[#333] mb-0.5">
                      {p.title}
                    </strong>
                    <span className="font-sans text-[13px] font-light text-[#696969] leading-relaxed">
                      {p.body || ''}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Bottom CTA */}
      <div className="px-5 pb-7 max-w-[520px] mx-auto w-full relative z-10">
        <button
          type="button"
          onClick={onStart}
          disabled={!config || loading}
          className={`
            w-full py-4 rounded-xl font-sans text-base font-semibold tracking-wide
            transition-all duration-300 ease-out
            ${
              config && !loading
                ? 'bg-[#8b6d4b] text-white hover:opacity-90 hover:-translate-y-0.5 active:translate-y-0 shadow-lg shadow-[#8b6d4b]/25 survey-cta-pulse'
                : 'bg-[#d0cbc6] text-white cursor-not-allowed'
            }
          `}
        >
          {ctaLabel}
        </button>

        {loading && !config && (
          <p className="text-center text-xs font-sans text-[#999] mt-2.5">
            설문 구성을 불러오는 중입니다.
          </p>
        )}

        <p className="text-center text-xs font-sans text-[#999] mt-3.5 leading-relaxed">
          선택하지 않은 공정의 질문은 자동으로 건너뜁니다.
          <br />
          필요한 것만 빠르게 답변하실 수 있습니다.
        </p>
      </div>

      {/* Pulse animation style */}
      <style>{`
        @keyframes ctaPulse {
          0%, 100% { box-shadow: 0 4px 14px rgba(139, 109, 75, 0.25); }
          50% { box-shadow: 0 4px 24px rgba(139, 109, 75, 0.4); }
        }
        .survey-cta-pulse {
          animation: ctaPulse 2.5s ease-in-out infinite;
        }
        .survey-cta-pulse:hover {
          animation: none;
        }
      `}</style>
    </div>
  );
}
