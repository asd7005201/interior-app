import { useEffect, useRef, useState } from 'react';
import { portfolioConfig } from '../config';

export default function Portfolio() {
  const sectionRef = useRef<HTMLElement>(null);
  const [isVisible, setIsVisible] = useState(false);
  const [activeFilter, setActiveFilter] = useState('전체');

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

  if (!portfolioConfig.heading) return null;

  const categories = ['전체', ...new Set(portfolioConfig.items.map((item) => item.category))];
  
  const filteredItems = activeFilter === '전체'
    ? portfolioConfig.items
    : portfolioConfig.items.filter((item) => item.category === activeFilter);

  return (
    <section
      id="portfolio"
      ref={sectionRef}
      className="py-24 md:py-32 bg-white"
    >
      <div className="max-w-[1400px] mx-auto px-6 lg:px-12">
        {/* Header */}
        <div
          className={`text-center mb-12 transition-all duration-1000 ${
            isVisible ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          <span className="text-xs font-medium tracking-[0.2em] uppercase text-[#8b6d4b] mb-4 block">
            {portfolioConfig.tag}
          </span>
          <h2 className="font-serif text-4xl md:text-5xl text-[#2c2c2c] mb-6">
            {portfolioConfig.heading}
          </h2>
          <p className="text-[#696969] max-w-2xl mx-auto text-lg leading-relaxed">
            {portfolioConfig.description}
          </p>
        </div>

        {/* Filter Tabs */}
        <div
          className={`flex justify-center gap-4 mb-12 flex-wrap transition-all duration-1000 delay-200 ${
            isVisible ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          {categories.map((category) => (
            <button
              key={category}
              onClick={() => setActiveFilter(category)}
              className={`px-6 py-3 rounded-full text-sm font-medium transition-all duration-300 ${
                activeFilter === category
                  ? 'bg-[#8b6d4b] text-white'
                  : 'bg-[#f5f5f5] text-[#696969] hover:bg-[#ebebeb]'
              }`}
            >
              {category}
            </button>
          ))}
        </div>

        {/* Portfolio Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {filteredItems.map((item, index) => (
            <div
              key={item.id}
              className={`group relative overflow-hidden rounded-2xl cursor-pointer transition-all duration-700 ${
                isVisible
                  ? 'opacity-100 translate-y-0'
                  : 'opacity-0 translate-y-12'
              }`}
              style={{ transitionDelay: `${index * 100}ms` }}
            >
              {/* Image */}
              <div className="relative h-80 md:h-96 overflow-hidden">
                <img
                  src={item.image}
                  alt={item.title}
                  className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-110"
                />
                {/* Overlay */}
                <div className="absolute inset-0 bg-gradient-to-t from-black/70 via-black/20 to-transparent opacity-60 group-hover:opacity-80 transition-opacity duration-500" />
              </div>

              {/* Content */}
              <div className="absolute bottom-0 left-0 right-0 p-8 text-white">
                <span className="inline-block px-3 py-1 bg-[#8b6d4b] rounded-full text-xs font-medium mb-4">
                  {item.category}
                </span>
                <h3 className="font-serif text-2xl md:text-3xl mb-3 group-hover:translate-y-0 transition-transform duration-500">
                  {item.title}
                </h3>
                <div className="flex gap-6 text-sm text-white/80 opacity-0 group-hover:opacity-100 transition-all duration-500 translate-y-4 group-hover:translate-y-0">
                  <span>{item.area}</span>
                  <span>{item.budget}</span>
                </div>
              </div>
            </div>
          ))}
        </div>

        {/* View All Button */}
        <div
          className={`text-center mt-12 transition-all duration-1000 delay-500 ${
            isVisible ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          <button className="inline-flex items-center gap-2 px-8 py-4 border-2 border-[#2c2c2c] text-[#2c2c2c] rounded-full font-medium hover:bg-[#2c2c2c] hover:text-white transition-all duration-300">
            {portfolioConfig.viewAllText}
            <svg
              className="w-5 h-5"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M17 8l4 4m0 0l-4 4m4-4H3"
              />
            </svg>
          </button>
        </div>
      </div>
    </section>
  );
}
