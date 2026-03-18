/* ── Skeleton loading state for SurveyLanding ────────────────────────────── */

function Bone({
  width,
  height,
  className = '',
}: {
  width: string;
  height: string;
  className?: string;
}) {
  return (
    <div
      className={`rounded-lg survey-skeleton-shimmer ${className}`}
      style={{ width, height }}
    />
  );
}

export default function SurveySkeleton() {
  return (
    <div className="min-h-screen flex flex-col bg-[#faf9f7] animate-fade-in">
      <div className="flex-1 flex flex-col items-center justify-center px-5 pt-12 pb-6">
        <div className="w-full max-w-[440px] mx-auto flex flex-col items-center">
          {/* Tag placeholder */}
          <Bone width="140px" height="28px" className="rounded-full mb-5" />

          {/* Title placeholder */}
          <Bone width="60%" height="28px" className="mb-2" />
          <Bone width="50%" height="28px" className="mb-5" />

          {/* Subtitle placeholder */}
          <Bone width="80%" height="16px" className="mb-2" />
          <Bone width="70%" height="16px" className="mb-8" />

          {/* Trust badges placeholder */}
          <div className="flex gap-2 justify-center mb-8">
            <Bone width="90px" height="28px" className="rounded-full" />
            <Bone width="100px" height="28px" className="rounded-full" />
            <Bone width="90px" height="28px" className="rounded-full" />
          </div>

          {/* Metrics placeholder */}
          <div className="flex gap-8 justify-center mb-8">
            {[0, 1, 2].map((i) => (
              <div key={i} className="flex flex-col items-center gap-1.5">
                <Bone width="48px" height="28px" />
                <Bone width="56px" height="12px" />
              </div>
            ))}
          </div>

          {/* Divider placeholder */}
          <Bone width="40px" height="2px" className="mb-8" />

          {/* Points card placeholder */}
          <div className="w-full rounded-2xl bg-white border border-[#f0ece8] p-6 mb-8">
            {[0, 1, 2].map((i) => (
              <div key={i} className={`flex gap-3.5 items-start ${i > 0 ? 'mt-5' : ''}`}>
                <Bone width="32px" height="32px" className="rounded-xl shrink-0" />
                <div className="flex-1">
                  <Bone width="40%" height="14px" className="mb-2" />
                  <Bone width="100%" height="12px" className="mb-1" />
                  <Bone width="80%" height="12px" />
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* CTA placeholder */}
      <div className="px-5 pb-7 max-w-[520px] mx-auto w-full">
        <Bone width="100%" height="52px" className="rounded-xl mx-auto" />
        <div className="flex justify-center mt-3.5">
          <Bone width="220px" height="12px" />
        </div>
      </div>

      {/* Shimmer animation */}
      <style>{`
        .survey-skeleton-shimmer {
          background: linear-gradient(
            90deg,
            #e8e4e0 25%,
            #f0ece8 50%,
            #e8e4e0 75%
          );
          background-size: 200% 100%;
          animation: skeletonShimmer 1.5s ease-in-out infinite;
        }
        @keyframes skeletonShimmer {
          0% { background-position: 200% 0; }
          100% { background-position: -200% 0; }
        }
      `}</style>
    </div>
  );
}
