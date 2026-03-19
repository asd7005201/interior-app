// ─── Site ────────────────────────────────────────────────────────────────────

export interface SiteConfig {
  title: string;
  description: string;
  language: string;
}

export const siteConfig: SiteConfig = {
  title: "우아한 인테리어 | 무료 온라인 견적",
  description: "인테리어 전문가가 직접 설계하고 시공하는 우아한 인테리어. 3분 설문으로 무료 가견적을 받아보세요. 투명한 가격, 검증된 자재, 책임 시공.",
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
    { label: "무료 견적", href: "#quote" },
    { label: "시공 사례", href: "#portfolio" },
    { label: "상담 신청", href: "#contact" },
  ],
  socialLinks: [
    { icon: "Instagram", label: "Instagram", href: "https://instagram.com" },
    { icon: "Facebook", label: "Facebook", href: "https://facebook.com" },
  ],
  searchPlaceholder: "궁금한 내용을 검색하세요",
  cartEmptyText: "아직 선택한 자재가 없습니다",
  cartCheckoutText: "견적 요청하기",
  continueShoppingText: "자재 더 둘러보기",
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
  tagline: "꿈꾸던 공간, 현실이 되는 순간",
  title: "당신의 집이\n작품이 됩니다",
  ctaPrimaryText: "3분 만에 무료 견적 받기",
  ctaPrimaryTarget: "#quote",
  ctaSecondaryText: "시공 사례 보기",
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
  tag: "우아한 인테리어 소개",
  heading: "살아보면 느끼는 차이,\n디테일이 다른 시공",
  bodyParagraphs: [
    "우아한 인테리어는 아파트, 빌라, 주택 등 주거 공간 전문으로 10년 넘게 고객의 라이프스타일에 맞춘 맞춤 설계를 해왔습니다.",
    "현장 실측부터 3D 디자인, 자재 선정, 시공, A/S까지 한 팀이 책임지는 원스톱 서비스. 견적 단계부터 자재 단가를 투명하게 공개해 추가 비용 걱정을 없앴습니다.",
  ],
  linkText: "서비스 자세히 보기",
  linkTarget: "#services",
  image1: "/images/about-studio.jpg",
  image2: "/images/about-living.jpg",
  stats: [
    { value: 520, suffix: "세대", label: "누적 시공 세대" },
    { value: 12, suffix: "년", label: "인테리어 전문 경력" },
    { value: 97, suffix: "%", label: "재시공 없는 완성률" },
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
  tag: "시공 프로세스",
  heading: "체계적인 5단계 시공",
  bodyParagraphs: [
    "현장 실측 → 3D 디자인 → 자재 확정 → 공정별 시공 → 최종 검수. 각 단계마다 고객 확인을 거쳐 만족스러운 결과를 만듭니다.",
    "담당 실장이 처음부터 끝까지 현장을 관리하여, 일정 지연과 품질 편차를 최소화합니다.",
  ],
  ctaText: "서비스 살펴보기",
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
  tag: "서비스 안내",
  heading: "공간별 맞춤 인테리어",
  description: "주방, 욕실, 침실 등 공간의 특성에 맞는 전문 설계와 시공으로 생활의 편리함과 아름다움을 동시에 드립니다.",
  services: [
    {
      id: 1,
      title: "주방 리모델링",
      description: "동선을 고려한 수납 설계와 내구성 높은 상판, 고급 수전으로 요리가 즐거워지는 주방을 만듭니다.",
      image: "/images/service-kitchen.jpg",
    },
    {
      id: 2,
      title: "욕실 리모델링",
      description: "방수 시공부터 타일, 수전, 욕조까지. 매일 사용하는 공간이니까 내구성과 디자인 모두 놓치지 않습니다.",
      image: "/images/service-bath.jpg",
    },
    {
      id: 3,
      title: "침실 & 거실",
      description: "조명 설계, 컬러 톤 조합, 맞춤 붙박이장으로 하루의 피로를 풀 수 있는 편안한 공간을 완성합니다.",
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
  tag: "무료 가견적",
  heading: "3분이면 충분합니다.\n무료로 견적을 받아보세요.",
  description: "간단한 설문만 작성하시면, 담당 실장이 검토 후 24시간 내 맞춤 견적서를 보내드립니다. 비용은 일절 들지 않습니다.",
  ctaText: "무료 견적 설문 시작하기",
  features: [
    "100% 무료, 부담 없는 가견적",
    "자재 단가까지 투명하게 공개",
    "공정별 상세 항목 안내",
    "설문 완료 후 24시간 내 회신",
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
  tag: "시공 사례",
  heading: "직접 확인하는 완성 품질",
  description: "평수별, 예산별 실제 시공 사례를 확인하고 나에게 맞는 스타일을 찾아보세요.",
  viewAllText: "전체 사례 보기",
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
      title: "1년 무상 A/S",
      description: "시공 완료 후 1년간 무상 보수. 하자 발생 시 48시간 내 방문 조치합니다.",
    },
    {
      icon: "Clock",
      title: "약속한 날짜에 완공",
      description: "공정표에 따라 일정을 관리하며, 지연 시 사전 안내드립니다.",
    },
    {
      icon: "Wallet",
      title: "추가 비용 제로",
      description: "계약 시 확정된 견적 그대로. 숨은 비용 없이 자재 단가까지 공개합니다.",
    },
    {
      icon: "Heart",
      title: "생활 맞춤 설계",
      description: "가족 구성, 생활 패턴, 취향을 반영한 실용적이고 아름다운 공간을 설계합니다.",
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
  tag: "자주 묻는 질문",
  heading: "궁금한 점, 미리 답해드립니다",
  ctaText: "더 궁금하신 점이 있으신가요? 편하게 문의해주세요.",
  ctaTarget: "#contact",
  faqs: [
    {
      id: 1,
      question: "온라인 견적은 어떻게 받나요?",
      answer: "홈페이지의 무료 견적 설문(약 3분)을 작성해주시면, 담당 실장이 내용을 검토한 뒤 24시간 내에 가견적서를 보내드립니다. 이후 현장 실측을 통해 최종 견적을 확정하며, 가견적 단계에서는 비용이 전혀 발생하지 않습니다.",
    },
    {
      id: 2,
      question: "시공 기간은 얼마나 걸리나요?",
      answer: "20평대 아파트 전체 시공 기준 보통 4~6주가 소요됩니다. 부분 시공(욕실, 주방 등)은 1~2주 내로 완료됩니다. 계약 시 공정표를 함께 드리며, 일정이 변경될 경우 사전에 안내드립니다.",
    },
    {
      id: 3,
      question: "자재는 직접 고를 수 있나요?",
      answer: "물론입니다. 바닥재, 타일, 수전, 도배지 등 주요 자재를 담당 디자이너와 함께 선택하실 수 있습니다. 자재 샘플을 직접 확인하신 후 결정하시면 되고, 선택에 따라 견적이 투명하게 조정됩니다.",
    },
    {
      id: 4,
      question: "A/S는 어떻게 되나요?",
      answer: "시공 완료 후 1년간 무상 A/S를 보장합니다. 하자 발생 시 접수 후 48시간 내 방문하여 조치하며, 무상 기간 이후에도 합리적인 비용으로 유상 A/S를 제공합니다. 시공한 팀이 직접 A/S를 담당하므로 빠르고 정확합니다.",
    },
    {
      id: 5,
      question: "대금 지급은 어떤 방식인가요?",
      answer: "계약금 30%, 중도금 40%, 잔금 30%로 나누어 공정 진행에 맞춰 지급합니다. 각 단계별 시공 결과를 확인하신 후 지급하시는 구조라 안심하실 수 있습니다. 카드 결제도 가능하며, 자세한 내용은 상담 시 안내드립니다.",
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
  heading: "편하게 연락주세요",
  description: "인테리어가 처음이셔도 괜찮습니다. 작은 궁금증이라도 부담 없이 남겨주시면, 경험 많은 실장이 친절하게 답변드리겠습니다.",
  locationLabel: "주소",
  location: "서울특별시 강남구 테헤란로 123 우아한빌딩 5F",
  emailLabel: "이메일",
  email: "hello@elegant-interior.co.kr",
  phoneLabel: "전화",
  phone: "02-1234-5678",
  formFields: {
    nameLabel: "이름",
    namePlaceholder: "홍길동",
    emailLabel: "이메일",
    emailPlaceholder: "example@email.com",
    messageLabel: "문의 내용",
    messagePlaceholder: "평수, 예산, 원하시는 스타일 등을 자유롭게 적어주세요",
  },
  submitText: "문의 보내기",
  submittingText: "보내는 중...",
  submittedText: "접수 완료",
  successMessage: "문의가 접수되었습니다. 영업일 기준 24시간 내에 답변드리겠습니다. 감사합니다.",
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
  brandDescription: "살아보면 느끼는 차이. 투명한 견적, 검증된 자재, 책임 시공으로 고객의 일상을 더 아름답게 만듭니다.",
  newsletterHeading: "인테리어 소식 받기",
  newsletterDescription: "시공 사례와 인테리어 팁을 이메일로 받아보세요.",
  newsletterPlaceholder: "이메일 주소를 입력하세요",
  newsletterButtonText: "구독하기",
  newsletterSuccessText: "구독이 완료되었습니다. 감사합니다!",
  linkGroups: [
    {
      title: "서비스",
      links: [
        { label: "주방 인테리어", href: "#services" },
        { label: "욕실 인테리어", href: "#services" },
        { label: "침실 인테리어", href: "#services" },
        { label: "무료 견적", href: "#quote" },
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
  tag: "추천 자재",
  heading: "검증된 인기 자재",
  description: "실제 시공에서 자주 사용되는 고품질 자재를 소개합니다. 자재 선택에 참고하세요.",
  viewAllText: "전체 자재 보기",
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
  tag: "인테리어 가이드",
  heading: "알면 도움 되는 인테리어 팁",
  viewAllText: "전체 글 보기",
  readMoreText: "자세히 읽기",
  posts: [
    {
      id: 1,
      title: "2026 인테리어 트렌드: 자연 소재와 따뜻한 미니멀리즘",
      date: "2026.01.15",
      image: "/images/hero-interior.jpg",
      excerpt: "올해 주목할 인테리어 트렌드를 정리했습니다. 우드톤, 자연광 활용, 곡선 가구 등 편안한 공간 만들기의 핵심을 알려드립니다.",
    },
    {
      id: 2,
      title: "20평대 아파트, 넓어 보이게 만드는 7가지 시공 팁",
      date: "2025.11.20",
      image: "/images/about-living.jpg",
      excerpt: "작은 평수도 설계만 잘하면 넓고 쾌적하게 살 수 있습니다. 전문가가 알려주는 공간 활용 노하우를 확인하세요.",
    },
    {
      id: 3,
      title: "주방 리모델링 전 반드시 알아야 할 5가지 체크리스트",
      date: "2025.09.05",
      image: "/images/service-kitchen.jpg",
      excerpt: "상판 소재 선택부터 수납 동선 설계까지, 후회 없는 주방 리모델링을 위한 핵심 포인트를 정리했습니다.",
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
      tag: "우리의 철학",
      heading: "공간에 가치를 더하다",
      paragraphs: [
        "우아한 인테리어는 단순히 예쁜 공간이 아니라, 그 안에서 살아가는 사람들의 일상이 더 편안해지는 것을 목표로 합니다.",
        "고객 한 분 한 분의 이야기를 듣고, 그 꿈을 현실로 만드는 것이 저희가 하는 일입니다.",
      ],
      quote: "좋은 공간은 좋은 삶의 시작입니다.",
      attribution: "— 우아한 인테리어",
      image: "/images/about-studio.jpg",
      backgroundColor: "#8b6d4b",
      textColor: "#ffffff",
    },
  ],
};
