import type { SurveyConfig } from '@/survey/types';

interface SurveyLandingProps {
  config: SurveyConfig | null;
  loading: boolean;
  onStart: () => void;
}

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

export default function SurveyLanding({ config, loading, onStart }: SurveyLandingProps) {
  const hero = uiHero(config);
  const chips = uiChips(config);
  const points = uiPoints(config);

  return (
    <div className="min-h-screen flex flex-col animate-fade-in">
      {/* Top section */}
      <div className="flex-1 flex flex-col items-center justify-center px-5 pt-10 pb-6 text-center">
        {/* Brand tagline */}
        <p className="font-sans text-sm font-semibold tracking-[0.12em] text-[#8b6d4b] mb-6">
          이상이 일상이 되는 시간
        </p>

        {/* Hero title */}
        <h1 className="font-serif text-[26px] leading-[1.35] tracking-tight mb-3 max-w-[360px] text-[#333]">
          {hero.title || config?.settings?.hero_title || '실무형 가견적, 빠르게 확인하세요'}
        </h1>

        {/* Hero subtitle */}
        {(hero.body || config?.settings?.hero_subtitle) && (
          <p className="font-sans text-sm font-light text-[#696969] max-w-[360px] mx-auto mb-6 leading-[1.7]">
            {hero.body || config?.settings?.hero_subtitle || ''}
          </p>
        )}

        {/* Chips */}
        {chips.length > 0 && (
          <div className="flex gap-2 justify-center flex-wrap mb-8">
            {chips.map((c, i) => (
              <span
                key={i}
                className="inline-flex items-center rounded-full bg-[#f5f3f0] px-3 py-1.5 text-xs font-sans font-medium text-[#696969]"
              >
                {c.title}
              </span>
            ))}
          </div>
        )}

        {/* Divider */}
        <div className="w-10 h-0.5 bg-[#8b6d4b] opacity-40 mx-auto mb-7" />

        {/* Points */}
        {points.length > 0 && (
          <div className="grid gap-4 w-full max-w-[400px] mx-auto mb-7 text-left">
            {points.map((p, i) => (
              <div key={i} className="flex gap-3.5 items-start">
                <div className="w-8 h-8 rounded-[10px] bg-[#8b6d4b]/10 text-[#8b6d4b] flex items-center justify-center shrink-0 font-sans font-bold text-sm">
                  {i + 1}
                </div>
                <div>
                  <strong className="block font-sans text-sm font-semibold text-[#333] mb-0.5">
                    {p.title}
                  </strong>
                  <span className="font-sans text-sm font-light text-[#696969] leading-relaxed">
                    {p.body || ''}
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Trust metrics */}
        <div className="flex gap-6 justify-center my-6">
          <div className="text-center">
            <div className="font-sans text-xl font-bold leading-tight text-[#8b6d4b]">4~6</div>
            <div className="font-sans text-[11px] text-[#999] mt-0.5">평균 소요 분</div>
          </div>
          <div className="text-center">
            <div className="font-sans text-xl font-bold leading-tight text-[#8b6d4b]">100%</div>
            <div className="font-sans text-[11px] text-[#999] mt-0.5">모바일 최적화</div>
          </div>
          <div className="text-center">
            <div className="font-sans text-xl font-bold leading-tight text-[#8b6d4b]">0</div>
            <div className="font-sans text-[11px] text-[#999] mt-0.5">불필요한 질문</div>
          </div>
        </div>
      </div>

      {/* Bottom CTA */}
      <div className="px-5 pb-6 max-w-[520px] mx-auto w-full">
        <button
          type="button"
          onClick={onStart}
          disabled={!config || loading}
          className={`
            w-full py-4 rounded-lg font-sans text-base font-semibold tracking-wide
            transition-all duration-300 ease-out
            ${
              config && !loading
                ? 'bg-[#8b6d4b] text-white hover:opacity-85 hover:-translate-y-0.5 active:translate-y-0 shadow-lg shadow-[#8b6d4b]/20'
                : 'bg-[#d0cbc6] text-white cursor-not-allowed'
            }
          `}
        >
          {(hero as Record<string, string>).cta_label || '가견적 시작하기'}
        </button>

        {loading && !config && (
          <p className="text-center text-xs font-sans text-[#999] mt-2.5">
            설문 구성을 불러오는 중입니다.
          </p>
        )}

        <p className="text-center text-xs font-sans text-[#999] mt-3 leading-relaxed">
          선택하지 않은 공정의 질문은 자동으로 건너뜁니다.
          <br />
          필요한 것만 빠르게 답변하실 수 있습니다.
        </p>
      </div>
    </div>
  );
}
