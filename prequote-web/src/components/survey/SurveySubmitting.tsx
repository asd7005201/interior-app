import { useEffect, useState, useRef } from 'react';

/* ── Stage definitions ───────────────────────────────────────────────────── */

const STAGES = [
  {
    target: 30,
    duration: 1500,
    text: '입력하신 내용을 정리하고 있습니다...',
  },
  {
    target: 70,
    duration: 2000,
    text: '공사 범위에 맞는 견적을 계산하고 있습니다...',
  },
  {
    target: 95,
    duration: 8000, // stays near 95 until parent resolves
    text: '거의 완료되었습니다. 잠시만 기다려주세요.',
  },
];

/* ── Component ───────────────────────────────────────────────────────────── */

export default function SurveySubmitting() {
  const [progress, setProgress] = useState(0);
  const [stageIdx, setStageIdx] = useState(0);
  const [fadingOut, setFadingOut] = useState(false);
  const rafRef = useRef<number>(0);
  const startRef = useRef(performance.now());
  const baseRef = useRef(0);

  useEffect(() => {
    const stage = STAGES[stageIdx];
    if (!stage) return;

    const from = baseRef.current;
    const to = stage.target;
    const duration = stage.duration;
    startRef.current = performance.now();

    const tick = (now: number) => {
      const elapsed = now - startRef.current;
      const t = Math.min(elapsed / duration, 1);
      // Ease-out quad
      const eased = 1 - (1 - t) * (1 - t);
      const current = from + (to - from) * eased;
      setProgress(current);

      if (t < 1) {
        rafRef.current = requestAnimationFrame(tick);
      } else if (stageIdx < STAGES.length - 1) {
        // Transition to next stage
        baseRef.current = to;
        setFadingOut(true);
        setTimeout(() => {
          setStageIdx((s) => s + 1);
          setFadingOut(false);
        }, 300);
      }
      // Last stage: just stay at 95
    };

    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [stageIdx]);

  const stage = STAGES[stageIdx];

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-6 bg-[#faf9f7]">
      <div className="w-full max-w-[360px] text-center">
        {/* Serif title */}
        <h2 className="font-serif text-xl text-[#333] mb-8">
          견적을 준비하고 있습니다
        </h2>

        {/* Progress bar */}
        <div className="w-full h-2 rounded-full bg-[#ece8e4] overflow-hidden mb-6">
          <div
            className="h-full rounded-full transition-[width] duration-100 ease-linear"
            style={{
              width: `${progress}%`,
              background: 'linear-gradient(90deg, #8b6d4b, #b9a79e)',
            }}
          />
        </div>

        {/* Percentage */}
        <div className="font-sans text-sm font-semibold text-[#8b6d4b] mb-4">
          {Math.round(progress)}%
        </div>

        {/* Stage text with fade */}
        <p
          className={`
            font-sans text-sm font-light text-[#696969] leading-relaxed
            transition-opacity duration-300
            ${fadingOut ? 'opacity-0' : 'opacity-100'}
          `}
        >
          {stage?.text}
        </p>
      </div>
    </div>
  );
}
