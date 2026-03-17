interface ProgressBarProps {
  current: number;
  total: number;
}

export default function ProgressBar({ current, total }: ProgressBarProps) {
  const pct = Math.round(((current + 1) / total) * 100);

  return (
    <div className="fixed top-0 left-0 right-0 z-50">
      <div className="h-1 bg-[#f5f3f0]">
        <div
          className="h-full bg-[#8b6d4b] transition-all duration-500 ease-out"
          style={{ width: `${pct}%` }}
        />
      </div>
      <div className="flex items-center justify-center py-2">
        <span className="text-xs font-sans font-medium tracking-wide text-[#696969]">
          {current + 1} / {total}
        </span>
      </div>
    </div>
  );
}
