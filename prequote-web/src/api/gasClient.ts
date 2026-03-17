const GAS_API_URL = import.meta.env.VITE_GAS_API_URL || '';

export async function gasCall<T = unknown>(action: string, params: Record<string, unknown> = {}): Promise<T> {
  if (!GAS_API_URL) {
    console.warn('[gasClient] VITE_GAS_API_URL not set, returning mock data');
    return {} as T;
  }
  const res = await fetch(GAS_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ action, ...params }),
    redirect: 'follow',
  });
  if (!res.ok) throw new Error(`API error: ${res.status}`);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'Unknown API error');
  return json.data as T;
}

// Public APIs
export const loadSurveyConfig = () => gasCall('loadSurveyConfig');
export const submitSurvey = (payload: Record<string, unknown>) => gasCall('submitSurvey', { payload });
export const getResultData = (id: string, token: string) => gasCall('getResultData', { id, token });
export const getResultRecommendations = (id: string, token: string) => gasCall('getResultRecommendations', { id, token });

// Admin APIs
export const adminLogin = (password: string) => gasCall('adminLogin', { password });
export const adminLogout = (credential: string) => gasCall('adminLogout', { credential });
export const adminGetDashboard = (credential: string) => gasCall('adminGetDashboard', { credential });
export const adminGetRequestDetail = (credential: string, requestId: string) => gasCall('adminGetRequestDetail', { credential, requestId });
export const adminSaveRequestReview = (credential: string, requestId: string, payload: Record<string, unknown>) => gasCall('adminSaveRequestReview', { credential, requestId, payload });
export const adminUpdateRequestStatus = (credential: string, requestId: string, newStatus: string, note?: string) => gasCall('adminUpdateRequestStatus', { credential, requestId, newStatus, note });
export const adminAddNote = (credential: string, requestId: string, noteText: string, options?: Record<string, unknown>) => gasCall('adminAddNote', { credential, requestId, noteText, options });
export const adminSearchMaterials = (credential: string, query: string, options?: Record<string, unknown>) => gasCall('adminSearchMaterials', { credential, query, options });
export const adminBulkUpdateRequests = (credential: string, payload: Record<string, unknown>) => gasCall('adminBulkUpdateRequests', { credential, payload });
export const adminCreateQuoteDraft = (credential: string, requestId: string, options?: Record<string, unknown>) => gasCall('adminCreateQuoteDraft', { credential, requestId, options });
export const adminSyncCaches = (credential: string) => gasCall('adminSyncCaches', { credential });
export const adminRecalculateRequest = (credential: string, requestId: string, payload?: Record<string, unknown>) => gasCall('adminRecalculateRequest', { credential, requestId, payload });
