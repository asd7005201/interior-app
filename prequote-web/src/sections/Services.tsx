import { useEffect, useRef, useState } from 'react';
import { servicesConfig } from '../config';

export default function Services() {
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

  if (!servicesConfig.heading) return null;

  return (
    <section
      id="services"
      ref={sectionRef}
      className="py-24 md:py-32 bg-[#fafafa]"
    >
      <div className="max-w-[1400px] mx-auto px-6 lg:px-12">
        {/* Header */}
        <div
          className={`text-center mb-16 transition-all duration-1000 ${
            isVisible ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          <span className="text-xs font-medium tracking-[0.2em] uppercase text-[#8b6d4b] mb-4 block">
            {servicesConfig.tag}
          </span>
          <h2 className="font-serif text-4xl md:text-5xl text-[#2c2c2c] mb-6">
            {servicesConfig.heading}
          </h2>
          <p className="text-[#696969] max-w-2xl mx-auto text-lg leading-relaxed">
            {servicesConfig.description}
          </p>
        </div>

        {/* Services Grid */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          {servicesConfig.services.map((service, index) => (
            <div
              key={service.id}
              className={`group bg-white rounded-2xl overflow-hidden shadow-sm hover:shadow-xl transition-all duration-500 ${
                isVisible
                  ? 'opacity-100 translate-y-0'
                  : 'opacity-0 translate-y-12'
              }`}
              style={{ transitionDelay: `${index * 150}ms` }}
            >
              {/* Image */}
              <div className="relative h-64 overflow-hidden">
                <img
                  src={service.image}
                  alt={service.title}
                  className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-110"
                />
                <div className="absolute inset-0 bg-gradient-to-t from-black/30 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-500" />
              </div>

              {/* Content */}
              <div className="p-8">
                <h3 className="font-serif text-2xl text-[#2c2c2c] mb-4 group-hover:text-[#8b6d4b] transition-colors duration-300">
                  {service.title}
                </h3>
                <p className="text-[#696969] leading-relaxed">
                  {service.description}
                </p>
              </div>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}
