interface SectionBreakProps {
  message: string;
  subtext?: string;
}

export default function SectionBreak({ message, subtext }: SectionBreakProps) {
  return (
    <div className="flex flex-col items-center justify-center py-10 px-5 text-center bg-[#faf9f7] section-break-fade-in">
      {/* Accent line */}
      <div className="w-8 h-0.5 bg-[#8b6d4b] opacity-40 mb-5" />

      {/* Sparkle / checkmark icon */}
      <div className="w-10 h-10 rounded-full bg-[#8b6d4b]/[0.08] flex items-center justify-center mb-4">
        <svg
          width="20"
          height="20"
          viewBox="0 0 20 20"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
          className="text-[#8b6d4b]"
        >
          <path
            d="M10 2L11.5 7.5L17 6L12.5 10L17 14L11.5 12.5L10 18L8.5 12.5L3 14L7.5 10L3 6L8.5 7.5L10 2Z"
            fill="currentColor"
            opacity="0.6"
          />
          <path
            d="M6.5 13.5L8.5 9.5L4.5 11.5L8.5 9.5L6.5 5.5"
            stroke="currentColor"
            strokeWidth="0"
          />
        </svg>
      </div>

      {/* Message */}
      <p className="font-serif text-lg text-[#333] leading-snug mb-2 max-w-[320px]">
        {message}
      </p>

      {/* Subtext */}
      {subtext && (
        <p className="font-sans text-sm font-light text-[#999] leading-relaxed max-w-[300px]">
          {subtext}
        </p>
      )}

      {/* Fade-in animation */}
      <style>{`
        .section-break-fade-in {
          animation: sectionBreakIn 0.5s ease-out forwards;
        }
        @keyframes sectionBreakIn {
          from {
            opacity: 0;
            transform: translateY(12px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
      `}</style>
    </div>
  );
}
