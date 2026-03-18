import { useMemo } from 'react';

interface ProgressBarProps {
  current: number;
  total: number;
  isBranchIntro?: boolean;
}

/**
 * Weighted progress: early steps feel faster (front-loaded weight curve).
 * Returns a percentage 0-100.
 */
function weightedProgress(current: number, total: number): number {
  if (total <= 1) return 100;
  // Each step i gets weight = total - i (earlier steps are heavier)
  const weights: number[] = [];
  for (let i = 0; i < total; i++) {
    weights.push(total - i);
  }
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  const doneWeight = weights.slice(0, current + 1).reduce((a, b) => a + b, 0);
  return Math.min(Math.round((doneWeight / totalWeight) * 100), 100);
}

/**
 * Estimate remaining minutes based on remaining step count.
 * Rough average ~18s per step.
 */
function estimateMinutes(remaining: number): number {
  const avgSecondsPerStep = 18;
  return Math.max(1, Math.ceil((remaining * avgSecondsPerStep) / 60));
}

export default function ProgressBar({
  current,
  total,
  isBranchIntro,
}: ProgressBarProps) {
  const pct = useMemo(
    () => weightedProgress(current, total),
    [current, total],
  );

  const remaining = total - current - 1;
  const minutes = estimateMinutes(remaining);

  return (
    <div className="fixed top-0 left-0 right-0 z-50 bg-white/80 backdrop-blur-sm">
      {/* Thin progress track */}
      <div className="h-1 bg-[#f0ece8]">
        <div
          className="relative h-full rounded-r-full transition-all duration-700 ease-out"
          style={{
            width: `${pct}%`,
            background: 'linear-gradient(90deg, #8b6d4b, #b9a79e)',
          }}
        >
          {/* Glowing leading dot */}
          <div
            className="absolute right-0 top-1/2 -translate-y-1/2 h-2.5 w-2.5 rounded-full"
            style={{
              background: '#8b6d4b',
              boxShadow: '0 0 8px 2px rgba(139,109,75,0.45)',
            }}
          />
        </div>
      </div>

      {/* Info row */}
      <div className="flex items-center justify-between px-4 py-1.5">
        <span className="text-[11px] font-sans font-medium tracking-wide text-[#8b6d4b]">
          {isBranchIntro ? '시작' : `${pct}%`}
        </span>
        {!isBranchIntro && remaining > 0 && (
          <span className="text-[11px] font-sans font-light text-[#999]">
            약 {minutes}분 남음
          </span>
        )}
      </div>
    </div>
  );
}
