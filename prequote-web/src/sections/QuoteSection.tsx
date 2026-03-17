import { useEffect, useRef, useState } from 'react';
import { Link } from 'react-router-dom';
import { quoteConfig } from '../config';
import { Check, Calculator, FileText, List, Eye, ArrowRight } from 'lucide-react';

export default function QuoteSection() {
  const sectionRef = useRef<HTMLElement>(null);
  const [isVisible, setIsVisible] = useState(false);

  useEffect(() => {
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (entry.isIntersecting) {
          setIsVisible(true);
          observer.disconnect();
        }
      },
      { threshold: 0.1 }
    );

    if (sectionRef.current) {
      observer.observe(sectionRef.current);
    }

    return () => observer.disconnect();
  }, []);

  if (!quoteConfig.heading) return null;

  const featureIcons = [Calculator, List, FileText, Eye];

  return (
    <section
      id="quote"
      ref={sectionRef}
      className="py-24 md:py-32 bg-[#f5f3f0] relative overflow-hidden"
    >
      {/* Background Pattern */}
      <div
        className="absolute inset-0 opacity-5"
        style={{
          backgroundImage: `url(${quoteConfig.backgroundImage})`,
          backgroundSize: 'cover',
          backgroundPosition: 'center',
        }}
      />

      <div className="max-w-[1400px] mx-auto px-6 lg:px-12 relative z-10">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-16 items-center">
          {/* Left Content */}
          <div
            className={`transition-all duration-1000 ${
              isVisible ? 'opacity-100 translate-x-0' : 'opacity-0 -translate-x-12'
            }`}
          >
            <span className="text-xs font-medium tracking-[0.2em] uppercase text-[#8b6d4b] mb-4 block">
              {quoteConfig.tag}
            </span>
            <h2 className="font-serif text-4xl md:text-5xl text-[#2c2c2c] mb-6 whitespace-pre-line">
              {quoteConfig.heading}
            </h2>
            <p className="text-[#696969] text-lg leading-relaxed mb-10">
              {quoteConfig.description}
            </p>

            {/* Features */}
            <div className="space-y-4">
              {quoteConfig.features.map((feature, index) => {
                const Icon = featureIcons[index] || Check;
                return (
                  <div
                    key={index}
                    className={`flex items-center gap-4 transition-all duration-700 ${
                      isVisible ? 'opacity-100 translate-x-0' : 'opacity-0 -translate-x-8'
                    }`}
                    style={{ transitionDelay: `${index * 100 + 300}ms` }}
                  >
                    <div className="w-10 h-10 rounded-full bg-[#8b6d4b]/10 flex items-center justify-center">
                      <Icon className="w-5 h-5 text-[#8b6d4b]" />
                    </div>
                    <span className="text-[#2c2c2c] font-medium">{feature}</span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Right CTA */}
          <div
            className={`transition-all duration-1000 delay-300 ${
              isVisible ? 'opacity-100 translate-x-0' : 'opacity-0 translate-x-12'
            }`}
          >
            <div className="bg-white rounded-3xl p-8 md:p-10 shadow-xl flex flex-col items-center justify-center text-center min-h-[320px]">
              <h3 className="font-serif text-2xl md:text-3xl text-[#2c2c2c] mb-4">
                간단한 설문으로<br />맞춤 견적을 받아보세요
              </h3>
              <p className="text-[#696969] mb-8 max-w-sm">
                3분이면 충분합니다. 공간 정보를 입력하시면 맞춤형 인테리어 견적을 무료로 제공해드립니다.
              </p>
              <Link
                to="/survey"
                className="inline-flex items-center gap-3 px-10 py-5 bg-[#8b6d4b] text-white rounded-2xl font-medium text-lg hover:bg-[#6d5539] transition-all duration-300 hover:scale-105 hover:shadow-lg"
              >
                무료 가견적 받기
                <ArrowRight className="w-5 h-5" />
              </Link>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}
