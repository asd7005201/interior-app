import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import { ChevronLeft, Plus, X, CheckCircle2 } from 'lucide-react';

import type {
  SurveyConfig,
  SurveyOption,
  Answers,
  Step,
  SerializedFile,
  SubmitResult,
} from '@/survey/types';
import { buildSteps, applyBundle, getOptionsForQuestion } from '@/survey/stepBuilder';
import { filterTradesBySpaces } from '@/survey/visibilityEngine';
import { loadSurveyConfig, submitSurvey } from '@/api/gasClient';

import ProgressBar from '@/components/survey/ProgressBar';
import SingleSelectCard from '@/components/survey/SingleSelectCard';
import MultiSelectCard from '@/components/survey/MultiSelectCard';
import SurveyLanding from '@/components/survey/SurveyLanding';
import SurveySubmitting from '@/components/survey/SurveySubmitting';
import SurveySkeleton from '@/components/survey/SurveySkeleton';
import SectionBreak from '@/components/survey/SectionBreak';

// ─── Constants ────────────────────────────────────────────────────────────────

const CONFIG_CACHE_KEY_PREFIX = 'pq_survey_cfg:';
const RESULT_BOOTSTRAP_PREFIX = 'pq_result_bootstrap:';
const AUTO_ADVANCE_DELAY = 220;

// ─── Cache helpers ────────────────────────────────────────────────────────────

function getCacheKey(settings: Record<string, string>): string {
  const appV = (settings.app_version || '').trim();
  const surveyV = (settings.survey_version || '').trim();
  return CONFIG_CACHE_KEY_PREFIX + appV + ':' + surveyV;
}

function readCachedConfig(): SurveyConfig | null {
  try {
    // Try all matching keys since we don't know version yet
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key?.startsWith(CONFIG_CACHE_KEY_PREFIX)) {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : null;
      }
    }
  } catch {
    /* ignore */
  }
  return null;
}

function writeCachedConfig(cfg: SurveyConfig): void {
  try {
    const key = getCacheKey(cfg.settings || {});
    localStorage.setItem(key, JSON.stringify(cfg));
  } catch {
    /* ignore */
  }
}

function storeResultBootstrap(
  result: SubmitResult,
  answers: Answers,
  contact: Record<string, string>,
  _config: SurveyConfig,
): void {
  if (!result?.request_id || !result?.share_token) return;
  const pt = String(answers.Q000_PROJECT_TYPE || '').toUpperCase();
  const bootstrap = {
    _bootstrap: true,
    request: {
      request_id: result.request_id,
      customer_name: contact.name || '',
      project_type: pt,
      created_at: new Date().toISOString(),
      status: 'NEW',
    },
    estimate: result.estimate || { min: 0, max: 0, flags: [] },
    flags: result.flags || [],
    recommendations: [],
    answers: [],
    answer_groups: [],
    notice_text: '',
    recommendation_status: 'PENDING',
    recommendation_count: 0,
  };
  try {
    sessionStorage.setItem(
      RESULT_BOOTSTRAP_PREFIX + result.request_id + ':' + result.share_token,
      JSON.stringify(bootstrap),
    );
  } catch {
    /* ignore */
  }
}

// ─── Validation ───────────────────────────────────────────────────────────────

function checkRequired(questions: Step['questions'], answers: Answers): boolean {
  return questions.every((q) => {
    // is_required comes as boolean or "Y" string from GAS
    const required =
      q.is_required === true || (q.is_required as unknown) === 'Y';
    if (!required) return true;
    const v = answers[q.question_code];
    if (v === undefined || v === null) return false;
    if (typeof v === 'string') return v.trim().length > 0;
    if (Array.isArray(v)) return v.length > 0;
    return true;
  });
}

// ─── Component ────────────────────────────────────────────────────────────────

type Screen = 'landing' | 'survey' | 'submitting' | 'done';

export default function SurveyPage() {
  const navigate = useNavigate();

  // Core state
  const [config, setConfig] = useState<SurveyConfig | null>(readCachedConfig);
  const [loading, setLoading] = useState(!readCachedConfig());
  const [error, setError] = useState<string | null>(null);
  const [screen, setScreen] = useState<Screen>('landing');
  const [step, setStep] = useState(0);
  const [answers, setAnswers] = useState<Answers>({});
  const [files, setFiles] = useState<File[]>([]);
  const [direction, setDirection] = useState<'next' | 'prev'>('next');
  const [transitioning, setTransitioning] = useState(false);

  // Track previous step count for clamping
  const prevStepCountRef = useRef(0);
  const autoAdvanceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // ── Load config ──────────────────────────────────────────────────────────

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const data = (await loadSurveyConfig()) as SurveyConfig;
        if (cancelled || !data) return;
        setConfig(data);
        writeCachedConfig(data);
        setLoading(false);
      } catch (err) {
        if (cancelled) return;
        setLoading(false);
        if (!config) {
          setError((err as Error).message || '설문 구성을 불러오지 못했습니다.');
        }
      }
    }

    if (config) {
      // Background refresh
      setLoading(false);
      setTimeout(() => { load(); }, 1200);
    } else {
      load();
    }

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Build steps ──────────────────────────────────────────────────────────

  const steps: Step[] = useMemo(() => {
    if (!config) return [];
    return buildSteps(
      answers,
      config.questions || [],
      config.options || [],
      config.scopeMap || [],
      config.bundles || [],
    );
  }, [config, answers]);

  // Clamp step when steps change
  useEffect(() => {
    const newCount = steps.length;
    if (newCount !== prevStepCountRef.current) {
      if (step >= newCount) {
        setStep(Math.max(0, newCount - 1));
      }
      prevStepCountRef.current = newCount;
    }
  }, [steps, step]);

  // Cleanup auto-advance timer
  useEffect(() => {
    return () => {
      if (autoAdvanceTimerRef.current) clearTimeout(autoAdvanceTimerRef.current);
    };
  }, []);

  // ── Helpers ──────────────────────────────────────────────────────────────

  const getOpts = useCallback(
    (questionId: string): SurveyOption[] => {
      if (!config) return [];
      return getOptionsForQuestion(config.options || [], questionId);
    },
    [config],
  );

  const currentStep = steps[step] || null;
  const isLastStep = step >= steps.length - 1;
  const allOptional = currentStep
    ? currentStep.questions.every(
        (q) => q.is_required !== true && (q.is_required as unknown) !== 'Y',
      )
    : false;
  const canProceed = currentStep ? checkRequired(currentStep.questions, answers) : false;

  // Detect branch intro (only entry questions)
  const isBranchIntro = currentStep
    ? currentStep.questions.every((q) => {
        const code = (q.question_code || '').trim();
        return code === 'Q000_PROJECT_TYPE' || code === 'Q000_FLOW_MODE';
      })
    : false;

  // ── Actions ──────────────────────────────────────────────────────────────

  const handleStart = useCallback(() => {
    if (!config) return;
    setScreen('survey');
    setStep(0);
    setAnswers({});
    setFiles([]);
    prevStepCountRef.current = 0;
  }, [config]);

  const handleSingleSelect = useCallback(
    (questionCode: string, value: string) => {
      setAnswers((prev) => {
        let next = { ...prev, [questionCode]: value };

        // Apply bundle defaults for preset selections
        if (
          questionCode.startsWith('Q000_PRESET') &&
          questionCode !== 'Q000_PRESET_TONE' &&
          config
        ) {
          next = applyBundle(value, next, config.bundles || []);
        }

        return next;
      });

      // Auto-advance for single-question steps
      const shouldAutoAdvance =
        config?.settings?.auto_advance_single_select !== 'false';

      if (shouldAutoAdvance && currentStep?.questions.length === 1) {
        if (autoAdvanceTimerRef.current) clearTimeout(autoAdvanceTimerRef.current);
        autoAdvanceTimerRef.current = setTimeout(() => {
          setDirection('next');
          setTransitioning(true);
          setTimeout(() => {
            setStep((s) => {
              // Re-check steps length at execution time
              // (steps may have changed due to answer)
              return s < steps.length - 1 ? s + 1 : s;
            });
            setTransitioning(false);
            window.scrollTo(0, 0);
          }, 200);
        }, AUTO_ADVANCE_DELAY);
      }
    },
    [config, currentStep, steps.length],
  );

  const handleMultiToggle = useCallback(
    (questionCode: string, value: string) => {
      setAnswers((prev) => {
        const arr = Array.isArray(prev[questionCode])
          ? [...(prev[questionCode] as string[])]
          : [];
        const idx = arr.indexOf(value);
        if (idx >= 0) arr.splice(idx, 1);
        else arr.push(value);

        const next = { ...prev, [questionCode]: arr };

        // If spaces changed, clean up trades
        if (questionCode === 'R010_SPACES') {
          const trades = prev.R012_TRADES;
          if (Array.isArray(trades)) {
            const spaces = arr;
            const hasAll = spaces.includes('ALL');
            if (!hasAll) {
              const TRADE_TO_SPACE: Record<string, string> = {
                R_BATH: 'BATH',
                R_KITCHEN: 'KITCHEN',
              };
              next.R012_TRADES = trades.filter((t) => {
                const req = TRADE_TO_SPACE[t];
                return !req || spaces.includes(req);
              });
            }
          }
        }

        return next;
      });
    },
    [],
  );

  const handleTextInput = useCallback((questionCode: string, value: string) => {
    setAnswers((prev) => ({ ...prev, [questionCode]: value }));
  }, []);

  const handleAddFiles = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const fileList = e.target.files;
      if (!fileList) return;
      const maxCount = Math.max(
        Number(config?.settings?.file_upload_max_count || 10),
        1,
      );
      const maxBytes =
        Math.max(Number(config?.settings?.file_upload_max_mb || 20), 1) *
        1024 *
        1024;

      setFiles((prev) => {
        const next = [...prev];
        for (let i = 0; i < fileList.length; i++) {
          if (next.length >= maxCount) break;
          if (fileList[i].size > maxBytes) continue;
          next.push(fileList[i]);
        }
        return next;
      });
      e.target.value = '';
    },
    [config],
  );

  const handleRemoveFile = useCallback((index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const handleNext = useCallback(() => {
    if (step < steps.length - 1) {
      setDirection('next');
      setTransitioning(true);
      setTimeout(() => {
        setStep((s) => s + 1);
        setTransitioning(false);
        window.scrollTo(0, 0);
      }, 200);
    }
  }, [step, steps.length]);

  const handlePrev = useCallback(() => {
    if (step > 0) {
      setDirection('prev');
      setTransitioning(true);
      setTimeout(() => {
        setStep((s) => s - 1);
        setTransitioning(false);
        window.scrollTo(0, 0);
      }, 200);
    } else {
      // Undo entry selections or go back to landing
      setAnswers((prev) => {
        const fm = prev.Q000_FLOW_MODE;
        const pt = prev.Q000_PROJECT_TYPE;
        if (fm) {
          const next = { ...prev };
          delete next.Q000_FLOW_MODE;
          return next;
        }
        if (pt) {
          const next = { ...prev };
          delete next.Q000_PROJECT_TYPE;
          return next;
        }
        return prev;
      });

      // If no project type, go to landing
      if (!answers.Q000_PROJECT_TYPE) {
        setScreen('landing');
      }
      prevStepCountRef.current = 0;
    }
  }, [step, answers]);

  const handleSubmit = useCallback(async () => {
    if (!config) return;
    setScreen('submitting');

    const a = { ...answers };
    const contact = {
      name: String(a.Q900_NAME || ''),
      phone: String(a.Q901_PHONE || ''),
      email: String(a.Q902_EMAIL || ''),
      method: String(a.Q903_CONTACT_METHOD || 'KAKAO'),
      note: String(a.Q904_NOTE || ''),
      address: String(a.Q906_ADDRESS || ''),
    };

    const serializedFiles: SerializedFile[] = files.map((f, idx) => ({
      file_name: f.name || '',
      file_mime: f.type || '',
      file_size: f.size || 0,
      upload_type: 'CLIENT_METADATA_ONLY' as const,
      sort_order: idx + 1,
    }));

    try {
      const result = (await submitSurvey({
        answers: a,
        contact,
        files: serializedFiles,
      })) as SubmitResult;

      storeResultBootstrap(result, a, contact, config);
      setScreen('done');

      // Auto-redirect to result page
      setTimeout(() => {
        navigate(
          `/result?id=${encodeURIComponent(result.request_id)}&token=${encodeURIComponent(result.share_token)}`,
        );
      }, 1500);
    } catch (err) {
      setScreen('survey');
      alert('오류: ' + ((err as Error).message || '제출 중 오류가 발생했습니다.'));
    }
  }, [config, answers, files, navigate]);

  // ── Render: Error ────────────────────────────────────────────────────────

  if (error && !config) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center px-5 text-center">
        <h2 className="font-serif text-2xl text-[#333] mb-3">오류</h2>
        <p className="font-sans text-sm text-[#696969] mb-6">{error}</p>
        <button
          type="button"
          onClick={() => window.location.reload()}
          className="px-6 py-2.5 rounded-lg border border-[#e8e4e0] font-sans text-sm font-medium text-[#333] hover:bg-[#f5f3f0] transition-colors"
        >
          다시 시도
        </button>
      </div>
    );
  }

  // ── Render: Landing ──────────────────────────────────────────────────────

  if (screen === 'landing') {
    if (loading && !config) {
      return <SurveySkeleton />;
    }
    return (
      <SurveyLanding config={config} loading={loading} onStart={handleStart} />
    );
  }

  // ── Render: Submitting ───────────────────────────────────────────────────

  if (screen === 'submitting') {
    return <SurveySubmitting />;
  }

  // ── Render: Done ─────────────────────────────────────────────────────────

  if (screen === 'done') {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center px-5 text-center animate-fade-in">
        <div className="w-16 h-16 rounded-full bg-emerald-50 flex items-center justify-center mb-5">
          <CheckCircle2 className="h-8 w-8 text-emerald-500" />
        </div>
        <h2 className="font-serif text-2xl text-[#333] mb-3">견적 결과가 준비되었습니다</h2>
        <p className="font-sans text-sm text-[#696969]">전문 컨설턴트가 확인 후 상세 안내드리겠습니다.</p>
      </div>
    );
  }

  // ── Render: Survey ───────────────────────────────────────────────────────

  if (!currentStep) return null;

  const ctaLabel = isLastStep
    ? isBranchIntro
      ? '선택 후 계속'
      : '결과 확인하기'
    : '다음';

  return (
    <div className="min-h-screen bg-[#faf9f7] pb-28">
      {/* Progress bar */}
      <ProgressBar current={step} total={steps.length} isBranchIntro={isBranchIntro} />

      {/* Back button */}
      <button
        type="button"
        onClick={handlePrev}
        className="fixed top-10 left-3 z-[101] w-10 h-10 rounded-full bg-white shadow-md border-0 flex items-center justify-center hover:bg-[#f5f3f0] active:bg-[#ece8e4] transition-colors"
      >
        <ChevronLeft className="h-5 w-5 text-[#696969]" />
      </button>

      {/* Step content */}
      <div
        key={step}
        className={`pt-16 max-w-[520px] mx-auto px-5 ${
          transitioning
            ? 'opacity-0 transition-opacity duration-200'
            : direction === 'next'
              ? 'animate-[slideInRight_300ms_ease-out]'
              : 'animate-[slideInLeft_300ms_ease-out]'
        }`}
      >
        {/* Section break */}
        {currentStep.sectionBreakMessage && (
          <SectionBreak
            message={currentStep.sectionBreakMessage}
            subtext={currentStep.sectionBreakSubtext}
          />
        )}

        {/* Step header */}
        <div className="mb-6 animate-fade-in">
          <h2 className="font-serif text-xl text-[#333] leading-snug">
            {currentStep.title}
          </h2>
          {currentStep.desc && (
            <p className="font-sans text-sm font-light text-[#696969] mt-1 leading-relaxed">
              {currentStep.desc}
            </p>
          )}
        </div>

        {/* Questions */}
        {currentStep.questions.map((q, qi) => {
          const code = q.question_code;
          const type = q.question_type;
          const isMultiQuestionStep = currentStep.questions.length > 1;

          return (
            <div key={code} className="mb-6 animate-fade-up" style={{ animationDelay: `${qi * 60}ms` }}>
              {/* Separator for multi-question steps */}
              {qi > 0 && isMultiQuestionStep && (
                <div className="h-px bg-[#e8e4e0] my-5 mb-6" />
              )}

              {/* Sub-title for multi-question steps */}
              {isMultiQuestionStep && q.title && (
                <h3 className="font-sans text-sm font-semibold text-[#333] mb-1">
                  {q.title}
                </h3>
              )}
              {isMultiQuestionStep && q.description && (
                <p className="font-sans text-xs font-light text-[#696969] mb-3 leading-relaxed">
                  {q.description}
                </p>
              )}

              {/* SINGLE_SELECT */}
              {type === 'SINGLE_SELECT' && (() => {
                const opts = getOpts(q.question_id || code);
                const hasImages = opts.some((o) => !!o.option_image_url);
                return (
                <div className={hasImages ? 'grid grid-cols-2 gap-3' : 'grid gap-2.5'}>
                  {opts.map((opt) => (
                    <SingleSelectCard
                      key={opt.option_code}
                      option={opt}
                      selected={answers[code] === opt.option_code}
                      onSelect={(v) => handleSingleSelect(code, v)}
                    />
                  ))}
                  {(answers[code] === 'OTHER' || answers[code] === 'OTHER_TEXT') && (
                    <input
                      type="text"
                      placeholder="직접 입력해주세요"
                      value={String(answers[code + '_OTHER_TEXT'] || '')}
                      onChange={(e) => handleTextInput(code + '_OTHER_TEXT', e.target.value)}
                      className="w-full mt-2 px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all"
                    />
                  )}
                </div>
                );
              })()}

              {/* MULTI_SELECT */}
              {type === 'MULTI_SELECT' && (() => {
                let opts = getOpts(q.question_id || code);
                if (code === 'R012_TRADES') {
                  opts = filterTradesBySpaces(opts, answers);
                }
                const selected = Array.isArray(answers[code])
                  ? (answers[code] as string[])
                  : [];
                const hasImages = opts.some((o) => !!o.option_image_url);
                return (
                  <div className={hasImages ? 'grid grid-cols-2 gap-3' : 'grid gap-2.5'}>
                    {opts.map((opt) => (
                      <MultiSelectCard
                        key={opt.option_code}
                        option={opt}
                        selected={selected.includes(opt.option_code)}
                        onToggle={(v) => handleMultiToggle(code, v)}
                      />
                    ))}
                    {selected.includes('OTHER') && (
                      <input
                        type="text"
                        placeholder="직접 입력해주세요"
                        value={String(answers[code + '_OTHER_TEXT'] || '')}
                        onChange={(e) =>
                          handleTextInput(code + '_OTHER_TEXT', e.target.value)
                        }
                        className="w-full mt-2 px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all"
                      />
                    )}
                  </div>
                );
              })()}

              {/* SHORT_TEXT */}
              {type === 'SHORT_TEXT' && (
                <input
                  type="text"
                  placeholder={q.placeholder || ''}
                  value={String(answers[code] || '')}
                  onChange={(e) => handleTextInput(code, e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all"
                />
              )}

              {/* LONG_TEXT */}
              {type === 'LONG_TEXT' && (
                <textarea
                  placeholder={q.placeholder || ''}
                  value={String(answers[code] || '')}
                  onChange={(e) => handleTextInput(code, e.target.value)}
                  rows={4}
                  className="w-full px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all resize-none"
                />
              )}

              {/* PHONE */}
              {type === 'PHONE' && (
                <input
                  type="tel"
                  placeholder={q.placeholder || '010-0000-0000'}
                  value={String(answers[code] || '')}
                  onChange={(e) => handleTextInput(code, e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all"
                />
              )}

              {/* EMAIL */}
              {type === 'EMAIL' && (
                <input
                  type="email"
                  placeholder={q.placeholder || 'name@example.com'}
                  value={String(answers[code] || '')}
                  onChange={(e) => handleTextInput(code, e.target.value)}
                  className="w-full px-4 py-3 rounded-lg border border-[#e8e4e0] font-sans text-sm text-[#333] placeholder:text-[#b9a79e] focus:outline-none focus:border-[#8b6d4b] focus:ring-2 focus:ring-[#8b6d4b]/10 transition-all"
                />
              )}

              {/* FILE */}
              {type === 'FILE' && (
                <div>
                  <label
                    htmlFor={`file_${code}`}
                    className="flex flex-col items-center justify-center w-full py-7 px-5 rounded-xl border-2 border-dashed border-[#d0cbc6] bg-white cursor-pointer hover:border-[#8b6d4b] hover:bg-[#8b6d4b]/[0.02] transition-all"
                  >
                    <Plus className="h-7 w-7 text-[#b9a79e] mb-2" />
                    <span className="font-sans text-sm text-[#696969]">
                      탭하여 사진 또는 파일을 선택하세요
                    </span>
                    <input
                      type="file"
                      id={`file_${code}`}
                      multiple
                      accept="image/*,.pdf"
                      className="hidden"
                      onChange={handleAddFiles}
                    />
                  </label>
                  {files.length > 0 && (
                    <div className="mt-3 grid gap-2">
                      {files.map((f, fi) => (
                        <div
                          key={fi}
                          className="flex items-center gap-2 px-3 py-2 rounded-lg bg-[#f5f3f0] font-sans text-sm text-[#333]"
                        >
                          <span className="truncate flex-1">{f.name}</span>
                          <button
                            type="button"
                            onClick={() => handleRemoveFile(fi)}
                            className="shrink-0 p-1 text-red-400 hover:text-red-600 transition-colors"
                          >
                            <X className="h-4 w-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* Helper text */}
              {q.helper_text && (
                <p className="mt-1.5 font-sans text-[11px] text-[#999] leading-relaxed">
                  {q.helper_text}
                </p>
              )}
            </div>
          );
        })}
      </div>

      {/* Sticky CTA */}
      <div className="fixed bottom-0 left-0 right-0 z-50 bg-white/95 backdrop-blur-sm border-t border-[#f0ece8] py-4 px-5">
        <div className="max-w-[520px] mx-auto">
          {isLastStep ? (
            <button
              type="button"
              onClick={handleSubmit}
              disabled={!canProceed}
              className={`
                w-full py-3.5 rounded-lg font-sans text-sm font-semibold tracking-wide
                transition-all duration-300
                ${
                  canProceed
                    ? 'bg-[#8b6d4b] text-white hover:opacity-85 shadow-lg shadow-[#8b6d4b]/20'
                    : 'bg-[#d0cbc6] text-white cursor-not-allowed'
                }
              `}
            >
              {ctaLabel}
            </button>
          ) : (
            <button
              type="button"
              onClick={handleNext}
              disabled={!canProceed}
              className={`
                w-full py-3.5 rounded-lg font-sans text-sm font-semibold tracking-wide
                transition-all duration-300
                ${
                  canProceed
                    ? 'bg-[#8b6d4b] text-white hover:opacity-85 shadow-lg shadow-[#8b6d4b]/20'
                    : 'bg-[#d0cbc6] text-white cursor-not-allowed'
                }
              `}
            >
              {ctaLabel}
            </button>
          )}
          {!isLastStep && allOptional && (
            <button
              type="button"
              onClick={handleNext}
              className="block w-full text-center mt-2 font-sans text-xs text-[#999] hover:text-[#696969] transition-colors"
            >
              나중에 답변 가능
            </button>
          )}
          {/* Trust micro-copy */}
          {(() => {
            const hasContact = currentStep.questions.some((q) => q.question_code.startsWith('Q90'));
            const hasBudget = currentStep.questions.some((q) => q.question_code.includes('BUDGET'));
            const trustText = hasContact
              ? '개인정보는 견적 상담 목적으로만 사용됩니다'
              : hasBudget
                ? '정확한 견적을 위한 참고 정보입니다'
                : isLastStep
                  ? '제출 후 전문 컨설턴트가 직접 검토합니다'
                  : null;
            return trustText ? (
              <p className="text-center mt-2 font-sans text-[11px] text-[#999] leading-relaxed">
                {trustText}
              </p>
            ) : null;
          })()}
        </div>
      </div>
    </div>
  );
}
