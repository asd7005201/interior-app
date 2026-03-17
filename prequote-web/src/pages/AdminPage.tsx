import { useEffect, useState, useCallback, useRef, type ReactNode } from 'react';
import {
  adminLogin,
  adminGetDashboard,
  adminGetRequestDetail,
  adminUpdateRequestStatus,
  adminAddNote,
  adminCreateQuoteDraft,
} from '../api/gasClient';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import {
  Search,
  LogOut,
  RefreshCw,
  X,
  ChevronRight,
  Send,
  FileText,
  User,
  Phone,
  Mail,
  MapPin,
  Clock,
  CalendarCheck,
  AlertTriangle,
  Inbox,
  Eye,
  UserCheck,
  ArrowRightCircle,
  Package,
  MessageSquare,
  Lock,
  Loader2,
  ShieldCheck,
} from 'lucide-react';

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

interface QueueStats {
  new_requests?: number;
  in_review?: number;
  today_due?: number;
  overdue_reminder?: number;
  quote_waiting?: number;
}

interface RequestRow {
  request_id: string;
  customer_name?: string;
  contact_phone?: string;
  contact_email?: string;
  address_label?: string;
  project_type_label?: string;
  housing_type_label?: string;
  area_label?: string;
  scope_level_label?: string;
  status?: string;
  status_label?: string;
  review_status?: string;
  review_status_label?: string;
  work_state?: string;
  work_state_label?: string;
  assignee_name?: string;
  created_at?: string;
  estimate_min?: number;
  estimate_max?: number;
  auto_estimate_min?: number;
  auto_estimate_max?: number;
  estimate_source_label?: string;
  priority_label?: string;
  next_action_summary?: string;
  reminder_at?: string;
  duplicate_customer_count?: number;
  is_manual_override?: boolean;
  quote_draft_status?: string;
}

interface Dashboard {
  requests?: RequestRow[];
  queue?: QueueStats;
}

interface DetailFlag {
  severity?: string;
  severity_label?: string;
  label?: string;
  flag?: string;
}

interface DetailAnswerItem {
  question_label?: string;
  question_code?: string;
  answer_display_text?: string;
  answer_display_values?: string[];
}

interface DetailAnswerGroup {
  group_label?: string;
  group_key?: string;
  items?: DetailAnswerItem[];
}

interface DetailRecommendation {
  rec_id?: string;
  title?: string;
  material_id?: string;
  brand?: string;
  spec?: string;
  subtitle?: string;
  image_url?: string;
  reason_text?: string;
  price_hint_min?: number;
  price_hint_max?: number;
  tags?: string[];
  trade_code?: string;
  material_type?: string;
}

interface TimelineEvent {
  event_type?: string;
  event_label?: string;
  created_at?: string;
  note?: string;
  actor?: string;
}

interface RequestDetail {
  request?: RequestRow;
  estimate?: { min?: number; max?: number };
  auto_estimate?: { min?: number; max?: number };
  flags?: DetailFlag[];
  answer_groups?: DetailAnswerGroup[];
  recommendations?: DetailRecommendation[];
  timeline?: TimelineEvent[];
  review?: {
    review_status?: string;
    override_estimate_min?: number;
    override_estimate_max?: number;
    estimate_reason?: string;
    review_note?: string;
  };
  notes?: Array<{ note_text?: string; created_at?: string; created_by?: string }>;
}

/* ------------------------------------------------------------------ */
/*  Constants                                                          */
/* ------------------------------------------------------------------ */

const STORAGE_KEY = 'prequote_admin_token_v3';
const POLL_INTERVAL = 60_000;

const STATUS_LABELS: Record<string, string> = {
  NEW: '신규 접수',
  CONTACTED: '1차 연락',
  SCHEDULED: '상담 예정',
  CONVERTED: '실견적 전환',
  CLOSED: '종료',
};

const STATUS_COLORS: Record<string, string> = {
  NEW: 'bg-blue-100 text-blue-700 border-blue-200',
  CONTACTED: 'bg-amber-100 text-amber-700 border-amber-200',
  SCHEDULED: 'bg-purple-100 text-purple-700 border-purple-200',
  CONVERTED: 'bg-emerald-100 text-emerald-700 border-emerald-200',
  CLOSED: 'bg-gray-100 text-gray-500 border-gray-200',
};

const FILTER_TABS = [
  { key: '', label: '전체' },
  { key: 'NEW', label: '신규' },
  { key: 'CONTACTED', label: '검토중' },
  { key: 'SCHEDULED', label: '대기' },
  { key: 'CONVERTED,CLOSED', label: '완료' },
];

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

function num(v: unknown): number {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function formatMan(v: unknown): string {
  return Math.round(num(v) / 10000).toLocaleString('ko-KR');
}

function formatRange(minV: unknown, maxV: unknown): string {
  return `${formatMan(minV)} ~ ${formatMan(maxV)}만원`;
}

function formatDateTime(v: unknown): string {
  if (!v) return '-';
  const d = new Date(v as string);
  if (isNaN(d.getTime())) return String(v);
  return new Intl.DateTimeFormat('ko-KR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(d);
}

function formatDate(v: unknown): string {
  if (!v) return '-';
  const d = new Date(v as string);
  if (isNaN(d.getTime())) return String(v);
  return new Intl.DateTimeFormat('ko-KR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(d);
}

function loadToken(): string {
  try {
    return localStorage.getItem(STORAGE_KEY) || '';
  } catch {
    return '';
  }
}

function saveToken(token: string) {
  try {
    localStorage.setItem(STORAGE_KEY, token);
  } catch { /* noop */ }
}

function clearToken() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch { /* noop */ }
}

/* ------------------------------------------------------------------ */
/*  Component                                                          */
/* ------------------------------------------------------------------ */

export default function AdminPage() {
  const [credential, setCredential] = useState(() => loadToken());
  const [authenticated, setAuthenticated] = useState(false);
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);

  const [dashboard, setDashboard] = useState<Dashboard | null>(null);
  const [dashboardLoading, setDashboardLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterTab, setFilterTab] = useState('');

  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [detail, setDetail] = useState<RequestDetail | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);

  const [noteText, setNoteText] = useState('');
  const [noteLoading, setNoteLoading] = useState(false);
  const [statusChangeLoading, setStatusChangeLoading] = useState(false);

  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  /* ---------- Auth ---------- */

  const handleLogout = useCallback(() => {
    clearToken();
    setCredential('');
    setAuthenticated(false);
    setDashboard(null);
    setSelectedId(null);
    setDetail(null);
  }, []);

  const handleAuthError = useCallback(
    (err: unknown) => {
      const msg = err instanceof Error ? err.message : '';
      if (msg.includes('401') || msg.includes('Unauthorized') || msg.includes('expired') || msg.includes('invalid')) {
        handleLogout();
      }
    },
    [handleLogout],
  );

  const loadDashboard = useCallback(
    async (cred: string) => {
      setDashboardLoading(true);
      try {
        const res = (await adminGetDashboard(cred)) as Dashboard;
        setDashboard(res || {});
      } catch (err) {
        handleAuthError(err);
      } finally {
        setDashboardLoading(false);
      }
    },
    [handleAuthError],
  );

  const handleLogin = async () => {
    const pw = loginPassword.trim();
    if (!pw) {
      setLoginError('관리자 비밀번호를 입력해 주세요.');
      return;
    }
    setLoginLoading(true);
    setLoginError('');
    try {
      const res = (await adminLogin(pw)) as { token?: string };
      const token = res?.token || pw;
      saveToken(token);
      setCredential(token);
      setAuthenticated(true);
      await loadDashboard(token);
    } catch (err) {
      setLoginError(err instanceof Error ? err.message : '로그인에 실패했습니다.');
    } finally {
      setLoginLoading(false);
    }
  };

  const handleSessionReuse = async () => {
    const stored = loadToken();
    if (!stored) {
      setLoginError('저장된 세션이 없습니다.');
      return;
    }
    setLoginLoading(true);
    setLoginError('');
    try {
      await loadDashboard(stored);
      setCredential(stored);
      setAuthenticated(true);
    } catch {
      setLoginError('세션이 만료되었습니다. 다시 로그인해 주세요.');
    } finally {
      setLoginLoading(false);
    }
  };

  // auto-login with stored token
  useEffect(() => {
    const stored = loadToken();
    if (stored) {
      setLoginLoading(true);
      loadDashboard(stored)
        .then(() => {
          setCredential(stored);
          setAuthenticated(true);
        })
        .catch(() => {
          /* stay on login */
        })
        .finally(() => setLoginLoading(false));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // polling
  useEffect(() => {
    if (!authenticated || !credential) return;
    pollRef.current = setInterval(() => {
      loadDashboard(credential);
    }, POLL_INTERVAL);
    return () => {
      if (pollRef.current) clearInterval(pollRef.current);
    };
  }, [authenticated, credential, loadDashboard]);

  /* ---------- Detail ---------- */

  const openDetail = useCallback(
    async (requestId: string) => {
      setSelectedId(requestId);
      setDetailLoading(true);
      setDetail(null);
      setNoteText('');
      try {
        const res = (await adminGetRequestDetail(credential, requestId)) as RequestDetail;
        setDetail(res);
      } catch (err) {
        handleAuthError(err);
      } finally {
        setDetailLoading(false);
      }
    },
    [credential, handleAuthError],
  );

  const closeDetail = () => {
    setSelectedId(null);
    setDetail(null);
  };

  const handleStatusChange = async (newStatus: string) => {
    if (!selectedId || !detail) return;
    setStatusChangeLoading(true);
    try {
      await adminUpdateRequestStatus(credential, selectedId, newStatus);
      // refresh
      await openDetail(selectedId);
      await loadDashboard(credential);
    } catch (err) {
      handleAuthError(err);
    } finally {
      setStatusChangeLoading(false);
    }
  };

  const handleAddNote = async () => {
    if (!selectedId || !noteText.trim()) return;
    setNoteLoading(true);
    try {
      await adminAddNote(credential, selectedId, noteText.trim());
      setNoteText('');
      await openDetail(selectedId);
    } catch (err) {
      handleAuthError(err);
    } finally {
      setNoteLoading(false);
    }
  };

  const handleCreateQuote = async () => {
    if (!selectedId) return;
    try {
      await adminCreateQuoteDraft(credential, selectedId);
      await openDetail(selectedId);
      await loadDashboard(credential);
    } catch (err) {
      handleAuthError(err);
    }
  };

  /* ---------- Filtering ---------- */

  const rows = dashboard?.requests || [];
  const queue = dashboard?.queue || {};

  const filteredRows = rows.filter((row) => {
    // tab filter
    if (filterTab) {
      const tabs = filterTab.split(',');
      if (!tabs.includes(row.status || 'NEW')) return false;
    }
    // search
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const hay = [
        row.customer_name,
        row.contact_phone,
        row.address_label,
        row.project_type_label,
        row.request_id,
        row.assignee_name,
      ]
        .join(' ')
        .toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  /* ================================================================ */
  /*  LOGIN SCREEN                                                     */
  /* ================================================================ */

  if (!authenticated) {
    return (
      <div className="min-h-screen bg-[#f5f3f0] relative overflow-hidden">
        {/* Decorative grid overlay */}
        <div className="pointer-events-none fixed inset-0 opacity-[0.04]" style={{
          backgroundImage: 'linear-gradient(90deg, #8b6d4b 1px, transparent 1px), linear-gradient(#8b6d4b 1px, transparent 1px)',
          backgroundSize: '40px 40px',
          maskImage: 'linear-gradient(180deg, rgba(0,0,0,.5), transparent 80%)',
          WebkitMaskImage: 'linear-gradient(180deg, rgba(0,0,0,.5), transparent 80%)',
        }} />

        <div className="min-h-screen flex items-center justify-center px-5">
          <div className="w-full max-w-[460px] animate-fade-in">
            <div className="bg-white/95 backdrop-blur-xl rounded-[28px] border border-white/70 shadow-[0_22px_64px_rgba(31,28,23,.14)] p-8">
              <div className="flex gap-2 mb-5">
                <Badge variant="secondary" className="bg-[#8b6d4b]/10 text-[#8b6d4b] border-[#8b6d4b]/15 text-[11px]">
                  가견적 운영 콘솔
                </Badge>
                <Badge variant="secondary" className="bg-amber-100 text-amber-700 border-amber-200 text-[11px]">
                  관리자 전용
                </Badge>
              </div>
              <h1 className="font-serif text-2xl md:text-[28px] text-[#2c2c2c] mb-2 leading-tight">
                관리자 로그인
              </h1>
              <p className="text-sm text-[#696969] mb-6">
                관리자 수정 금액, 추천 자재, 운영 메모, 견적앱 초안 생성까지 같은 화면에서 처리할 수 있습니다.
              </p>

              <div className="space-y-2">
                <label className="block text-xs font-bold text-[#696969] tracking-wide">
                  관리자 비밀번호
                </label>
                <div className="relative">
                  <Lock className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-[#b9a79e]" />
                  <Input
                    type="password"
                    placeholder="비밀번호를 입력하세요"
                    value={loginPassword}
                    onChange={(e) => setLoginPassword(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleLogin()}
                    className="pl-10 h-11 rounded-xl"
                  />
                </div>
              </div>

              <div className="flex gap-2.5 mt-5">
                <Button
                  onClick={handleLogin}
                  disabled={loginLoading}
                  className="flex-1 h-11 rounded-xl bg-[#8b6d4b] hover:bg-[#7a5e3d] font-semibold"
                >
                  {loginLoading ? <Loader2 className="h-4 w-4 animate-spin" /> : '로그인'}
                </Button>
                <Button
                  variant="outline"
                  onClick={handleSessionReuse}
                  disabled={loginLoading}
                  className="h-11 rounded-xl"
                >
                  저장된 세션
                </Button>
              </div>

              {loginError && (
                <p className="mt-3 text-sm text-red-500 flex items-center gap-1.5">
                  <AlertTriangle className="h-3.5 w-3.5" />
                  {loginError}
                </p>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  /* ================================================================ */
  /*  DASHBOARD                                                        */
  /* ================================================================ */

  return (
    <div className="min-h-screen bg-[#f5f3f0] relative">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-[#faf9f5]/90 backdrop-blur-xl border-b border-[rgba(122,116,108,.16)]">
        <div className="mx-auto max-w-[1540px] flex items-center justify-between gap-4 px-5 py-3">
          <div>
            <h1 className="font-serif text-lg md:text-xl text-[#2c2c2c]">가견적 관리자</h1>
          </div>
          <div className="flex items-center gap-2.5">
            <Badge variant="secondary" className="text-[11px] bg-[rgba(122,116,108,.08)] text-[#696969]">
              <ShieldCheck className="h-3 w-3 mr-1" />
              {credential.startsWith('AS_') ? '세션 토큰' : '비밀번호 인증'}
            </Badge>
            <Button
              variant="outline"
              size="sm"
              onClick={() => loadDashboard(credential)}
              disabled={dashboardLoading}
              className="rounded-lg"
            >
              <RefreshCw className={`h-3.5 w-3.5 ${dashboardLoading ? 'animate-spin' : ''}`} />
              <span className="hidden sm:inline ml-1">새로고침</span>
            </Button>
            <Button variant="outline" size="sm" onClick={handleLogout} className="rounded-lg">
              <LogOut className="h-3.5 w-3.5" />
              <span className="hidden sm:inline ml-1">로그아웃</span>
            </Button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-[1540px] px-5 py-5 space-y-5">
        {/* ---- KPI Strip ---- */}
        <section className="grid grid-cols-2 md:grid-cols-4 gap-3 animate-fade-in">
          <KpiCard
            icon={<Inbox className="h-4 w-4" />}
            label="신규"
            value={queue.new_requests || 0}
            accent="from-blue-500 to-blue-600"
            sub="최초 검토 필요"
          />
          <KpiCard
            icon={<Eye className="h-4 w-4" />}
            label="검토필요"
            value={queue.in_review || 0}
            accent="from-amber-500 to-amber-600"
            sub="관리자 확인 대기"
          />
          <KpiCard
            icon={<CalendarCheck className="h-4 w-4" />}
            label="오늘 처리"
            value={queue.today_due || 0}
            accent="from-emerald-500 to-emerald-600"
            sub="오늘 액션 필요"
          />
          <KpiCard
            icon={<AlertTriangle className="h-4 w-4" />}
            label="오버듀"
            value={queue.overdue_reminder || 0}
            accent="from-red-500 to-red-600"
            sub="기한 초과"
          />
        </section>

        {/* ---- Filters ---- */}
        <section className="bg-white/90 rounded-2xl border border-[rgba(122,116,108,.14)] shadow-[0_12px_34px_rgba(33,31,27,.06)] p-4 animate-fade-in">
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2 flex-wrap">
              {FILTER_TABS.map((tab) => (
                <button
                  key={tab.key}
                  onClick={() => setFilterTab(tab.key)}
                  className={`rounded-full px-3.5 py-1.5 text-xs font-bold border transition-all duration-200 ${
                    filterTab === tab.key
                      ? 'bg-[#8b6d4b] border-[#8b6d4b] text-white shadow-md'
                      : 'bg-white border-[rgba(122,116,108,.14)] text-[#696969] hover:bg-[#f5f3f0]'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>
            <div className="relative w-full sm:w-auto sm:min-w-[280px]">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-[#b9a79e]" />
              <Input
                placeholder="이름, 연락처, 주소 검색"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="pl-10 h-9 rounded-lg text-sm"
              />
            </div>
          </div>
          <div className="mt-3 text-xs text-[#696969]">
            {filteredRows.length.toLocaleString('ko-KR')}건 표시
          </div>
        </section>

        {/* ---- Request Table / Cards ---- */}
        <section className="bg-white/90 rounded-2xl border border-[rgba(122,116,108,.14)] shadow-[0_12px_34px_rgba(33,31,27,.06)] overflow-hidden animate-fade-in">
          {dashboardLoading && !dashboard ? (
            <div className="p-8 space-y-3">
              {[1, 2, 3, 4].map((i) => (
                <Skeleton key={i} className="h-16 rounded-xl" />
              ))}
            </div>
          ) : filteredRows.length === 0 ? (
            <div className="py-16 text-center text-[#696969]">
              <Inbox className="h-10 w-10 mx-auto mb-3 text-[#b9a79e]" />
              <p className="text-sm">조건에 맞는 요청이 없습니다.</p>
            </div>
          ) : (
            <>
              {/* Desktop table */}
              <div className="hidden lg:block overflow-auto">
                <table className="w-full border-collapse">
                  <thead>
                    <tr className="border-b border-[rgba(122,116,108,.12)]">
                      {['이름', '연락처', '프로젝트', '면적', '상태', '담당자', '접수일'].map((h) => (
                        <th
                          key={h}
                          className="sticky top-0 bg-white/95 px-4 py-3 text-left text-[11px] font-bold text-[#696969] tracking-wider"
                        >
                          {h}
                        </th>
                      ))}
                      <th className="w-10" />
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row) => (
                      <tr
                        key={row.request_id}
                        onClick={() => openDetail(row.request_id)}
                        className={`border-b border-[rgba(122,116,108,.06)] cursor-pointer transition-colors duration-150 hover:bg-[#8b6d4b]/[0.03] ${
                          selectedId === row.request_id ? 'bg-[#8b6d4b]/[0.05]' : ''
                        }`}
                      >
                        <td className="px-4 py-3">
                          <div className="font-medium text-sm text-[#2c2c2c]">
                            {row.customer_name || '이름 미입력'}
                          </div>
                        </td>
                        <td className="px-4 py-3 text-sm text-[#696969]">
                          {row.contact_phone || '-'}
                        </td>
                        <td className="px-4 py-3 text-sm text-[#696969]">
                          {[row.project_type_label, row.housing_type_label].filter(Boolean).join(' / ') || '-'}
                        </td>
                        <td className="px-4 py-3 text-sm text-[#696969]">{row.area_label || '-'}</td>
                        <td className="px-4 py-3">
                          <StatusBadge status={row.status} />
                        </td>
                        <td className="px-4 py-3 text-sm text-[#696969]">
                          {row.assignee_name || '-'}
                        </td>
                        <td className="px-4 py-3 text-xs text-[#696969]">
                          {formatDate(row.created_at)}
                        </td>
                        <td className="px-2 py-3">
                          <ChevronRight className="h-4 w-4 text-[#b9a79e]" />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Mobile cards */}
              <div className="lg:hidden divide-y divide-[rgba(122,116,108,.08)]">
                {filteredRows.map((row) => (
                  <button
                    key={row.request_id}
                    onClick={() => openDetail(row.request_id)}
                    className="w-full text-left px-4 py-3.5 flex items-center gap-3 hover:bg-[#8b6d4b]/[0.03] transition-colors"
                  >
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="font-medium text-sm text-[#2c2c2c] truncate">
                          {row.customer_name || '이름 미입력'}
                        </span>
                        <StatusBadge status={row.status} />
                      </div>
                      <div className="mt-1 text-xs text-[#696969] truncate">
                        {[row.project_type_label, row.area_label, row.contact_phone]
                          .filter(Boolean)
                          .join(' / ')}
                      </div>
                      <div className="mt-0.5 text-[11px] text-[#b9a79e]">
                        {formatDate(row.created_at)}
                        {row.assignee_name ? ` / ${row.assignee_name}` : ''}
                      </div>
                    </div>
                    <ChevronRight className="h-4 w-4 shrink-0 text-[#b9a79e]" />
                  </button>
                ))}
              </div>
            </>
          )}
        </section>
      </main>

      {/* ============================================================ */}
      {/*  DETAIL PANEL (overlay)                                      */}
      {/* ============================================================ */}
      {selectedId && (
        <div
          className="fixed inset-0 z-50 bg-black/25 backdrop-blur-sm animate-fade-in"
          onClick={(e) => e.target === e.currentTarget && closeDetail()}
        >
          <div className="absolute inset-y-0 right-0 w-full max-w-[800px] bg-[#f7f6f1] shadow-[0_0_60px_rgba(24,22,20,.2)] flex flex-col animate-slide-left overflow-hidden">
            {/* Detail header */}
            <div className="sticky top-0 z-10 bg-[#f7f6f1]/95 backdrop-blur-xl border-b border-[rgba(122,116,108,.14)] px-6 py-4">
              <div className="flex items-start justify-between gap-4">
                <div className="min-w-0">
                  {detailLoading ? (
                    <Skeleton className="h-7 w-48 rounded-lg" />
                  ) : (
                    <>
                      <h2 className="font-serif text-xl md:text-2xl text-[#2c2c2c] truncate">
                        {detail?.request?.customer_name || '고객'}
                      </h2>
                      <div className="mt-1.5 flex items-center gap-2 flex-wrap text-xs text-[#696969]">
                        {detail?.request?.contact_phone && (
                          <span className="flex items-center gap-1">
                            <Phone className="h-3 w-3" />
                            {detail.request.contact_phone}
                          </span>
                        )}
                        {detail?.request?.contact_email && (
                          <span className="flex items-center gap-1">
                            <Mail className="h-3 w-3" />
                            {detail.request.contact_email}
                          </span>
                        )}
                        <StatusBadge status={detail?.request?.status} />
                      </div>
                    </>
                  )}
                </div>
                <button
                  onClick={closeDetail}
                  className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-white border border-[rgba(122,116,108,.14)] hover:bg-[#f5f3f0] transition-colors"
                >
                  <X className="h-4 w-4 text-[#696969]" />
                </button>
              </div>

              {/* Estimate bar */}
              {!detailLoading && detail?.estimate && (
                <div className="mt-4 rounded-2xl bg-gradient-to-r from-white to-[#f5f0e5] border border-[rgba(184,150,59,.2)] p-4">
                  <span className="text-[11px] font-bold text-[#696969] tracking-wide">예상 견적 범위</span>
                  <div className="font-serif text-2xl font-bold text-[#2c2c2c] mt-1">
                    {formatRange(detail.estimate.min || 0, detail.estimate.max || 0)}
                  </div>
                  {detail.auto_estimate && (
                    <p className="text-[11px] text-[#696969] mt-1">
                      자동 산출: {formatRange(detail.auto_estimate.min || 0, detail.auto_estimate.max || 0)}
                      {detail.review?.estimate_reason && ` / ${detail.review.estimate_reason}`}
                    </p>
                  )}
                </div>
              )}
            </div>

            {/* Detail scroll content */}
            <div className="flex-1 overflow-auto px-6 py-5 space-y-5">
              {detailLoading ? (
                <div className="space-y-4">
                  {[1, 2, 3].map((i) => (
                    <Skeleton key={i} className="h-32 rounded-2xl" />
                  ))}
                </div>
              ) : detail ? (
                <>
                  {/* Customer info */}
                  <DetailSection icon={<User className="h-4 w-4" />} title="고객 정보">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <MetaRow icon={<User className="h-3 w-3" />} label="이름" value={detail.request?.customer_name} />
                      <MetaRow icon={<Phone className="h-3 w-3" />} label="연락처" value={detail.request?.contact_phone} />
                      <MetaRow icon={<Mail className="h-3 w-3" />} label="이메일" value={detail.request?.contact_email} />
                      <MetaRow icon={<MapPin className="h-3 w-3" />} label="주소" value={detail.request?.address_label} />
                      <MetaRow label="프로젝트" value={[detail.request?.project_type_label, detail.request?.housing_type_label].filter(Boolean).join(' / ')} />
                      <MetaRow label="면적/범위" value={[detail.request?.area_label, detail.request?.scope_level_label].filter(Boolean).join(' / ')} />
                    </div>
                  </DetailSection>

                  {/* Answer summary */}
                  {(detail.answer_groups || []).length > 0 && (
                    <DetailSection icon={<FileText className="h-4 w-4" />} title="설문 응답 요약">
                      <div className="space-y-4">
                        {detail.answer_groups!.map((group, gi) => (
                          <div key={gi}>
                            <h4 className="text-xs font-bold text-[#696969] mb-2 tracking-wide">
                              {group.group_label || group.group_key || '응답'}
                            </h4>
                            <div className="space-y-1.5">
                              {(group.items || []).map((item, ai) => (
                                <div key={ai} className="flex items-start gap-2 text-sm">
                                  <span className="text-[11px] font-medium text-[#696969] min-w-[80px] pt-0.5">
                                    {item.question_label || item.question_code}
                                  </span>
                                  <span className="text-[#333]">
                                    {(item.answer_display_values || []).length > 1
                                      ? item.answer_display_values!.join(', ')
                                      : item.answer_display_text || '미입력'}
                                  </span>
                                </div>
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                    </DetailSection>
                  )}

                  {/* Flags */}
                  {(detail.flags || []).length > 0 && (
                    <DetailSection icon={<AlertTriangle className="h-4 w-4" />} title="체크 포인트">
                      <div className="space-y-2">
                        {detail.flags!.map((flag, i) => {
                          const sev = (flag.severity || '').toUpperCase();
                          const dotCls =
                            sev === 'HIGH'
                              ? 'bg-red-500'
                              : sev === 'MEDIUM'
                                ? 'bg-amber-500'
                                : 'bg-emerald-500';
                          return (
                            <div key={i} className="flex items-start gap-2.5 text-sm">
                              <span className={`mt-1.5 h-2 w-2 shrink-0 rounded-full ${dotCls}`} />
                              <div>
                                <span className="text-xs font-bold text-[#696969]">
                                  {flag.severity_label || flag.severity || '참고'}
                                </span>
                                <p className="text-[#333] mt-0.5">{flag.label || flag.flag}</p>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </DetailSection>
                  )}

                  {/* Timeline */}
                  {(detail.timeline || []).length > 0 && (
                    <DetailSection icon={<Clock className="h-4 w-4" />} title="타임라인">
                      <div className="relative pl-6 border-l-2 border-[rgba(122,116,108,.12)] space-y-4">
                        {detail.timeline!.map((evt, i) => (
                          <div key={i} className="relative">
                            <span className="absolute -left-[25px] top-1 h-3 w-3 rounded-full border-2 border-white bg-[#8b6d4b]" />
                            <div className="text-xs font-bold text-[#8b6d4b]">
                              {evt.event_label || evt.event_type}
                            </div>
                            <p className="text-xs text-[#696969] mt-0.5">{formatDateTime(evt.created_at)}</p>
                            {evt.note && <p className="text-sm text-[#333] mt-1">{evt.note}</p>}
                          </div>
                        ))}
                      </div>
                    </DetailSection>
                  )}

                  {/* Recommendations */}
                  {(detail.recommendations || []).length > 0 && (
                    <DetailSection icon={<Package className="h-4 w-4" />} title="추천 자재">
                      <div className="space-y-2.5">
                        {detail.recommendations!.map((item, i) => {
                          const summary = [item.brand, item.spec, item.subtitle].filter(Boolean).join(' / ');
                          return (
                            <div
                              key={item.rec_id || i}
                              className="flex gap-3 rounded-2xl border border-[rgba(122,116,108,.12)] bg-white p-3"
                            >
                              <div
                                className="h-16 w-16 shrink-0 rounded-xl bg-gradient-to-br from-[#ebe6dc] to-[#d8d1c4] bg-cover bg-center"
                                style={
                                  item.image_url
                                    ? { backgroundImage: `url('${item.image_url}')` }
                                    : undefined
                                }
                              />
                              <div className="min-w-0 flex-1">
                                <strong className="block text-sm font-bold font-serif text-[#2c2c2c]">
                                  {item.title || item.material_id || '추천 자재'}
                                </strong>
                                <p className="text-xs text-[#696969] truncate mt-0.5">
                                  {summary || '정보 준비 중'}
                                </p>
                                {item.reason_text && (
                                  <p className="text-xs text-[#696969] mt-1">{item.reason_text}</p>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </DetailSection>
                  )}

                  {/* Notes */}
                  {(detail.notes || []).length > 0 && (
                    <DetailSection icon={<MessageSquare className="h-4 w-4" />} title="운영 메모">
                      <div className="space-y-2">
                        {detail.notes!.map((note, i) => (
                          <div
                            key={i}
                            className="rounded-xl bg-white border border-[rgba(122,116,108,.1)] px-4 py-2.5 text-sm"
                          >
                            <p className="text-[#333]">{note.note_text}</p>
                            <p className="text-[11px] text-[#b9a79e] mt-1">
                              {note.created_by || 'ADMIN'} / {formatDateTime(note.created_at)}
                            </p>
                          </div>
                        ))}
                      </div>
                    </DetailSection>
                  )}

                  {/* Add note */}
                  <DetailSection icon={<Send className="h-4 w-4" />} title="메모 추가">
                    <div className="flex gap-2">
                      <Input
                        placeholder="운영 메모를 입력하세요"
                        value={noteText}
                        onChange={(e) => setNoteText(e.target.value)}
                        onKeyDown={(e) => e.key === 'Enter' && handleAddNote()}
                        className="flex-1 h-9 rounded-lg text-sm"
                      />
                      <Button
                        onClick={handleAddNote}
                        disabled={noteLoading || !noteText.trim()}
                        size="sm"
                        className="rounded-lg bg-[#8b6d4b] hover:bg-[#7a5e3d]"
                      >
                        {noteLoading ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Send className="h-3.5 w-3.5" />}
                      </Button>
                    </div>
                  </DetailSection>

                  {/* Actions */}
                  <div className="space-y-2.5 pb-4">
                    <h4 className="text-xs font-bold text-[#696969] tracking-wide flex items-center gap-1.5">
                      <ArrowRightCircle className="h-3.5 w-3.5 text-[#8b6d4b]" />
                      상태 변경
                    </h4>
                    <div className="flex flex-wrap gap-2">
                      {Object.entries(STATUS_LABELS).map(([key, label]) => (
                        <Button
                          key={key}
                          variant={detail?.request?.status === key ? 'default' : 'outline'}
                          size="sm"
                          disabled={statusChangeLoading || detail?.request?.status === key}
                          onClick={() => handleStatusChange(key)}
                          className={`rounded-lg text-xs ${
                            detail?.request?.status === key ? 'bg-[#8b6d4b] hover:bg-[#7a5e3d]' : ''
                          }`}
                        >
                          {label}
                        </Button>
                      ))}
                    </div>

                    <div className="flex flex-wrap gap-2 mt-3">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={handleCreateQuote}
                        className="rounded-lg text-xs"
                      >
                        <FileText className="h-3.5 w-3.5 mr-1" />
                        견적앱으로 전환
                      </Button>
                      <Button variant="outline" size="sm" className="rounded-lg text-xs">
                        <UserCheck className="h-3.5 w-3.5 mr-1" />
                        담당자 배정
                      </Button>
                    </div>
                  </div>
                </>
              ) : (
                <div className="py-16 text-center text-[#696969]">
                  <AlertTriangle className="h-8 w-8 mx-auto mb-2 text-[#b9a79e]" />
                  <p className="text-sm">상세 정보를 불러오지 못했습니다.</p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Sub-components                                                     */
/* ------------------------------------------------------------------ */

function KpiCard({
  icon,
  label,
  value,
  accent,
  sub,
}: {
  icon: ReactNode;
  label: string;
  value: number;
  accent: string;
  sub: string;
}) {
  return (
    <div
      className={`relative overflow-hidden rounded-2xl bg-gradient-to-br ${accent} p-4 text-white min-h-[88px] shadow-lg`}
    >
      <div className="absolute -right-4 -bottom-5 h-24 w-24 rounded-full bg-white/10" />
      <div className="flex items-center gap-1.5 text-white/80 text-xs font-medium">
        {icon}
        {label}
      </div>
      <div className="font-serif text-2xl font-bold mt-2">{value.toLocaleString('ko-KR')}</div>
      <p className="text-[11px] text-white/70 mt-1">{sub}</p>
    </div>
  );
}

function StatusBadge({ status }: { status?: string }) {
  const s = (status || 'NEW').toUpperCase();
  const colorCls = STATUS_COLORS[s] || STATUS_COLORS.NEW;
  const label = STATUS_LABELS[s] || s;
  return (
    <span
      className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[11px] font-bold ${colorCls}`}
    >
      {label}
    </span>
  );
}

function DetailSection({
  icon,
  title,
  children,
}: {
  icon: ReactNode;
  title: string;
  children: ReactNode;
}) {
  return (
    <section className="bg-white/80 rounded-2xl border border-[rgba(122,116,108,.1)] p-4">
      <h3 className="flex items-center gap-1.5 text-xs font-bold text-[#8b6d4b] tracking-wide mb-3">
        {icon}
        {title}
      </h3>
      {children}
    </section>
  );
}

function MetaRow({
  icon,
  label,
  value,
}: {
  icon?: ReactNode;
  label: string;
  value?: string;
}) {
  return (
    <div className="flex items-start gap-2 text-sm">
      {icon && <span className="mt-0.5 text-[#b9a79e]">{icon}</span>}
      <div>
        <span className="block text-[11px] font-bold text-[#696969] tracking-wide">{label}</span>
        <span className="text-[#333]">{value || '미입력'}</span>
      </div>
    </div>
  );
}
