// ─── Site ────────────────────────────────────────────────────────────────────

export interface SiteConfig {
  title: string;
  description: string;
  language: string;
}

export const siteConfig: SiteConfig = {
  title: "우아한 인테리어 | 견적 시스템",
  description: "전문 인테리어 디자인과 견적 서비스를 제공하는 우아한 인테리어입니다. 정확한 견적, 투명한 가격, 완벽한 시공을 약속드립니다.",
  language: "ko",
};

// ─── Navigation ──────────────────────────────────────────────────────────────

export interface MenuLink {
  label: string;
  href: string;
}

export interface SocialLink {
  icon: string;
  label: string;
  href: string;
}

export interface NavigationConfig {
  brandName: string;
  menuLinks: MenuLink[];
  socialLinks: SocialLink[];
  searchPlaceholder: string;
  cartEmptyText: string;
  cartCheckoutText: string;
  continueShoppingText: string;
  menuBackgroundImage: string;
}

export const navigationConfig: NavigationConfig = {
  brandName: "우아한 인테리어",
  menuLinks: [
    { label: "홈", href: "#hero" },
    { label: "서비스", href: "#services" },
    { label: "견적 문의", href: "#quote" },
    { label: "포트폴리오", href: "#portfolio" },
    { label: "문의하기", href: "#contact" },
  ],
  socialLinks: [
    { icon: "Instagram", label: "Instagram", href: "https://instagram.com" },
    { icon: "Facebook", label: "Facebook", href: "https://facebook.com" },
  ],
  searchPlaceholder: "검색어를 입력하세요...",
  cartEmptyText: "견적함이 비어있습니다",
  cartCheckoutText: "견적 요청하기",
  continueShoppingText: "계속 둘러보기",
  menuBackgroundImage: "/images/hero-interior.jpg",
};

// ─── Hero ────────────────────────────────────────────────────────────────────

export interface HeroConfig {
  tagline: string;
  title: string;
  ctaPrimaryText: string;
  ctaPrimaryTarget: string;
  ctaSecondaryText: string;
  ctaSecondaryTarget: string;
  backgroundImage: string;
}

export const heroConfig: HeroConfig = {
  tagline: "당신의 공간을 특별하게",
  title: "우아한 인테리어\n견적 시스템",
  ctaPrimaryText: "견적 문의하기",
  ctaPrimaryTarget: "#quote",
  ctaSecondaryText: "포트폴리오 보기",
  ctaSecondaryTarget: "#portfolio",
  backgroundImage: "/images/hero-interior.jpg",
};

// ─── SubHero ─────────────────────────────────────────────────────────────────

export interface Stat {
  value: number;
  suffix: string;
  label: string;
}

export interface SubHeroConfig {
  tag: string;
  heading: string;
  bodyParagraphs: string[];
  linkText: string;
  linkTarget: string;
  image1: string;
  image2: string;
  stats: Stat[];
}

export const subHeroConfig: SubHeroConfig = {
  tag: "About Us",
  heading: "공간에 영감을 불어넣는\n인테리어 전문가",
  bodyParagraphs: [
    "우아한 인테리어는 10년 이상의 경험을 바탕으로, 고객 한 분 한 분의 라이프스타일에 맞춘 맞춤형 인테리어 서비스를 제공합니다.",
    "정확한 견적 시스템과 투명한 가격 정책으로, 예산 내에서 최고의 품질을 약속드립니다. 디자인부터 시공까지 원스톱 서비스로 편리함을 드립니다.",
  ],
  linkText: "더 알아보기",
  linkTarget: "#services",
  image1: "/images/about-studio.jpg",
  image2: "/images/about-living.jpg",
  stats: [
    { value: 500, suffix: "+", label: "완료 프로젝트" },
    { value: 10, suffix: "+", label: "업계 경력(년)" },
    { value: 98, suffix: "%", label: "고객 만족도" },
  ],
};

// ─── Video Section ───────────────────────────────────────────────────────────

export interface VideoSectionConfig {
  tag: string;
  heading: string;
  bodyParagraphs: string[];
  ctaText: string;
  ctaTarget: string;
  backgroundImage: string;
}

export const videoSectionConfig: VideoSectionConfig = {
  tag: "Our Process",
  heading: "전문적인 시공 과정",
  bodyParagraphs: [
    "현장 실측부터 디자인, 자재 선정, 시공까지 체계적인 프로세스로 완벽한 공간을 만듭니다.",
    "각 단계별 전문가가 책임지고 진행하여 품질을 보장합니다.",
  ],
  ctaText: "더 알아보기",
  ctaTarget: "#services",
  backgroundImage: "/images/service-kitchen.jpg",
};

// ─── Services ────────────────────────────────────────────────────────────────

export interface Service {
  id: number;
  title: string;
  description: string;
  image: string;
}

export interface ServicesConfig {
  tag: string;
  heading: string;
  description: string;
  services: Service[];
}

export const servicesConfig: ServicesConfig = {
  tag: "Our Services",
  heading: "전문 인테리어 서비스",
  description: "주거 공간부터 상업 공간까지, 전문적인 디자인과 시공으로 완벽한 공간을 만들어드립니다.",
  services: [
    {
      id: 1,
      title: "주방 인테리어",
      description: "기능성과 미관을 동시에 잡은 맞춤형 주방 디자인. 수납 공간 최적화와 고급 자재로 특별한 주방을 선사합니다.",
      image: "/images/service-kitchen.jpg",
    },
    {
      id: 2,
      title: "욕실 인테리어",
      description: "스파 같은 편안함을 집 안에. 고급 욕조와 타일, 수전으로 특별한 힐링 공간을 디자인합니다.",
      image: "/images/service-bath.jpg",
    },
    {
      id: 3,
      title: "침실 인테리어",
      description: "편안한 수면을 위한 최적의 공간 설계. 조명, 컬러, 가구 배치로 완벽한 휴식 공간을 만듭니다.",
      image: "/images/service-bedroom.jpg",
    },
  ],
};

// ─── Quote Section ───────────────────────────────────────────────────────────

export interface QuoteConfig {
  tag: string;
  heading: string;
  description: string;
  ctaText: string;
  features: string[];
  backgroundImage: string;
}

export const quoteConfig: QuoteConfig = {
  tag: "Get a Quote",
  heading: "정확한 견적,\n투명한 가격",
  description: "간단한 정보 입력으로 정확한 견적을 받아보세요. 담당자가 확인 후 상세 견적서를 볂내드립니다.",
  ctaText: "견적 요청하기",
  features: [
    "실시간 견적 계산",
    "투명한 자재 단가",
    "상세 공정별 견적",
    "온라인 견적서 확인",
  ],
  backgroundImage: "/images/materials-samples.jpg",
};

// ─── Portfolio ───────────────────────────────────────────────────────────────

export interface PortfolioItem {
  id: number;
  title: string;
  category: string;
  image: string;
  area: string;
  budget: string;
}

export interface PortfolioConfig {
  tag: string;
  heading: string;
  description: string;
  viewAllText: string;
  items: PortfolioItem[];
}

export const portfolioConfig: PortfolioConfig = {
  tag: "Portfolio",
  heading: "완성된 프로젝트",
  description: "우아한 인테리어가 완성한 다양한 공간들을 만나보세요.",
  viewAllText: "전체 보기",
  items: [
    {
      id: 1,
      title: "강남 아파트 리모델링",
      category: "주거",
      image: "/images/hero-interior.jpg",
      area: "32평",
      budget: "5,000만원",
    },
    {
      id: 2,
      title: "홍대 카페 인테리어",
      category: "상업",
      image: "/images/service-kitchen.jpg",
      area: "20평",
      budget: "3,500만원",
    },
    {
      id: 3,
      title: "분당 주택 리모델링",
      category: "주거",
      image: "/images/service-bedroom.jpg",
      area: "45평",
      budget: "8,000만원",
    },
    {
      id: 4,
      title: "신사동 오피스텔",
      category: "주거",
      image: "/images/about-living.jpg",
      area: "15평",
      budget: "2,000만원",
    },
  ],
};

// ─── Features ────────────────────────────────────────────────────────────────

export interface Feature {
  icon: "Truck" | "ShieldCheck" | "Leaf" | "Heart" | "Clock" | "Wallet";
  title: string;
  description: string;
}

export interface FeaturesConfig {
  features: Feature[];
}

export const featuresConfig: FeaturesConfig = {
  features: [
    {
      icon: "ShieldCheck",
      title: "품질 보증",
      description: "모든 공정은 품질 보증 기간이 제공되며, A/S까지 책임집니다.",
    },
    {
      icon: "Clock",
      title: "정확한 일정",
      description: "약속한 일정에 맞춰 정확하게 시공을 완료합니다.",
    },
    {
      icon: "Wallet",
      title: "투명한 가격",
      description: "숨겨진 비용 없이 투명한 견적으로 예산을 관리하세요.",
    },
    {
      icon: "Heart",
      title: "맞춤 디자인",
      description: "고객의 라이프스타일에 맞는 맞춤형 디자인을 제공합니다.",
    },
  ],
};

// ─── FAQ ─────────────────────────────────────────────────────────────────────

export interface FaqItem {
  id: number;
  question: string;
  answer: string;
}

export interface FaqConfig {
  tag: string;
  heading: string;
  ctaText: string;
  ctaTarget: string;
  faqs: FaqItem[];
}

export const faqConfig: FaqConfig = {
  tag: "FAQ",
  heading: "자주 묻는 질문",
  ctaText: "더 궁금한 점이 있으신가요?",
  ctaTarget: "#contact",
  faqs: [
    {
      id: 1,
      question: "견적은 어떻게 받아볼 수 있나요?",
      answer: "홈페이지의 견적 문의 양식을 작성해주시면, 담당자가 연락드려 현장 방문 후 정확한 견적을 산출해드립니다. 온라인으로도 대략적인 견적 확인이 가능합니다.",
    },
    {
      id: 2,
      question: "시공 기간은 얼마나 걸리나요?",
      answer: "평수와 공정에 따라 다르지만, 일반적으로 20평대 아파트 기준 4~6주 정도 소요됩니다. 정확한 일정은 견적 상담 시 안내드립니다.",
    },
    {
      id: 3,
      question: "자재는 직접 선택할 수 있나요?",
      answer: "네, 가능합니다. 담당 디자이너와 상담을 통해 다양한 자재 옵션 중에서 선택하실 수 있으며, 자재에 따라 견적이 조정될 수 있습니다.",
    },
    {
      id: 4,
      question: "A/S는 어떻게 이루어지나요?",
      answer: "시공 완료 후 1년간 무상 A/S를 제공하며, 이후에도 유상으로 A/S가 가능합니다. 하자 발생 시 신속하게 처리해드립니다.",
    },
    {
      id: 5,
      question: "계약금과 대금 지급은 어떻게 하나요?",
      answer: "계약 시 30%의 계약금을 입금해주시며, 공정 진행 상황에 따라 중도금과 잔금을 나누어 지급합니다. 자세한 내용은 견적 상담 시 안내드립니다.",
    },
  ],
};

// ─── Contact ─────────────────────────────────────────────────────────────────

export interface FormFields {
  nameLabel: string;
  namePlaceholder: string;
  emailLabel: string;
  emailPlaceholder: string;
  messageLabel: string;
  messagePlaceholder: string;
}

export interface ContactConfig {
  heading: string;
  description: string;
  locationLabel: string;
  location: string;
  emailLabel: string;
  email: string;
  phoneLabel: string;
  phone: string;
  formFields: FormFields;
  submitText: string;
  submittingText: string;
  submittedText: string;
  successMessage: string;
  backgroundImage: string;
}

export const contactConfig: ContactConfig = {
  heading: "문의하기",
  description: "인테리어 관련 문의사항이 있으시면 언제든 연락주세요. 빠른 시일 내에 답변드리겠습니다.",
  locationLabel: "주소",
  location: "서울특별시 강남구 테헤란로 123 우아한빌딩 5F",
  emailLabel: "이메일",
  email: "hello@elegant-interior.co.kr",
  phoneLabel: "전화",
  phone: "02-1234-5678",
  formFields: {
    nameLabel: "이름",
    namePlaceholder: "이름을 입력하세요",
    emailLabel: "이메일",
    emailPlaceholder: "이메일을 입력하세요",
    messageLabel: "문의내용",
    messagePlaceholder: "문의하실 내용을 입력하세요",
  },
  submitText: "보내기",
  submittingText: "볂내는 중...",
  submittedText: "보내기 완료",
  successMessage: "문의가 성공적으로 접수되었습니다. 빠른 시일 내에 답변드리겠습니다.",
  backgroundImage: "/images/hero-interior.jpg",
};

// ─── Footer ──────────────────────────────────────────────────────────────────

export interface FooterLink {
  label: string;
  href: string;
}

export interface FooterLinkGroup {
  title: string;
  links: FooterLink[];
}

export interface FooterSocialLink {
  icon: string;
  label: string;
  href: string;
}

export interface FooterConfig {
  brandName: string;
  brandDescription: string;
  newsletterHeading: string;
  newsletterDescription: string;
  newsletterPlaceholder: string;
  newsletterButtonText: string;
  newsletterSuccessText: string;
  linkGroups: FooterLinkGroup[];
  legalLinks: FooterLink[];
  copyrightText: string;
  socialLinks: FooterSocialLink[];
}

export const footerConfig: FooterConfig = {
  brandName: "우아한 인테리어",
  brandDescription: "당신의 공간을 특별하게 만드는 인테리어 전문가. 정확한 견적, 투명한 가격, 완벽한 시공을 약속드립니다.",
  newsletterHeading: "뉴스레터 구독",
  newsletterDescription: "최신 프로젝트와 인테리어 트렌드를 받아보세요.",
  newsletterPlaceholder: "이메일 주소를 입력하세요",
  newsletterButtonText: "구독하기",
  newsletterSuccessText: "구독이 완료되었습니다!",
  linkGroups: [
    {
      title: "서비스",
      links: [
        { label: "주방 인테리어", href: "#services" },
        { label: "욕실 인테리어", href: "#services" },
        { label: "침실 인테리어", href: "#services" },
        { label: "견적 문의", href: "#quote" },
      ],
    },
    {
      title: "회사",
      links: [
        { label: "소개", href: "#about" },
        { label: "포트폴리오", href: "#portfolio" },
        { label: "문의하기", href: "#contact" },
      ],
    },
  ],
  legalLinks: [
    { label: "이용약관", href: "#" },
    { label: "개인정보처리방침", href: "#" },
  ],
  copyrightText: "© 2026 우아한 인테리어. All rights reserved.",
  socialLinks: [
    { icon: "Instagram", label: "Instagram", href: "https://instagram.com" },
    { icon: "Facebook", label: "Facebook", href: "https://facebook.com" },
  ],
};

// ─── Products (for cart functionality) ───────────────────────────────────────

export interface Product {
  id: number;
  name: string;
  price: number;
  category: string;
  image: string;
}

export interface ProductsConfig {
  tag: string;
  heading: string;
  description: string;
  viewAllText: string;
  addToCartText: string;
  addedToCartText: string;
  categories: string[];
  products: Product[];
}

export const productsConfig: ProductsConfig = {
  tag: "Materials",
  heading: "인기 자재",
  description: "우아한 인테리어에서 자주 사용하는 고품질 자재들을 소개합니다.",
  viewAllText: "전체 보기",
  addToCartText: "견적에 추가",
  addedToCartText: "추가되었습니다",
  categories: ["전체", "바닥재", "벽재", "주방", "욕실"],
  products: [
    {
      id: 1,
      name: "오크 원목 마루",
      price: 150000,
      category: "바닥재",
      image: "/images/materials-samples.jpg",
    },
    {
      id: 2,
      name: "대리석 타일",
      price: 80000,
      category: "벽재",
      image: "/images/service-bath.jpg",
    },
    {
      id: 3,
      name: "원목 주방 상판",
      price: 450000,
      category: "주방",
      image: "/images/service-kitchen.jpg",
    },
  ],
};

// ─── Blog ────────────────────────────────────────────────────────────────────

export interface BlogPost {
  id: number;
  title: string;
  date: string;
  image: string;
  excerpt: string;
}

export interface BlogConfig {
  tag: string;
  heading: string;
  viewAllText: string;
  readMoreText: string;
  posts: BlogPost[];
}

export const blogConfig: BlogConfig = {
  tag: "Journal",
  heading: "인테리어 이야기",
  viewAllText: "전체 보기",
  readMoreText: "자세히 보기",
  posts: [
    {
      id: 1,
      title: "2024년 인테리어 트렌드: 따뜻한 미니멀리즘",
      date: "2024.03.01",
      image: "/images/hero-interior.jpg",
      excerpt: "올해의 인테리어 트렌드를 알아보고, 집 안에 자연스럽게 녹아드는 디자인을 만나보세요.",
    },
    {
      id: 2,
      title: "작은 공간을 넓게 만드는 5가지 팁",
      date: "2024.02.15",
      image: "/images/about-living.jpg",
      excerpt: "원룸이나 소형 아파트도 넓고 쾌적하게! 공간 활용의 비결을 공개합니다.",
    },
    {
      id: 3,
      title: "주방 리모델링 가이드",
      date: "2024.01.20",
      image: "/images/service-kitchen.jpg",
      excerpt: "기능성과 디자인을 동시에 잡는 주방 리모델링의 모든 것.",
    },
  ],
};

// ─── About ───────────────────────────────────────────────────────────────────

export interface AboutSection {
  tag: string;
  heading: string;
  paragraphs: string[];
  quote: string;
  attribution: string;
  image: string;
  backgroundColor: string;
  textColor: string;
}

export interface AboutConfig {
  sections: AboutSection[];
}

export const aboutConfig: AboutConfig = {
  sections: [
    {
      tag: "Our Philosophy",
      heading: "공간에 가치를 더하다",
      paragraphs: [
        "우아한 인테리어는 단순히 예쁜 공간을 만드는 것을 넘어, 그곳에 사는 사람들의 삶의 질을 높이는 것을 목표로 합니다.",
        "모든 프로젝트는 고객의 꿈과 우리의 전문성이 만나 완성되는 특별한 여정입니다.",
      ],
      quote: "좋은 공간은 좋은 삶의 시작입니다.",
      attribution: "— 우아한 인테리어",
      image: "/images/about-studio.jpg",
      backgroundColor: "#8b6d4b",
      textColor: "#ffffff",
    },
  ],
};
