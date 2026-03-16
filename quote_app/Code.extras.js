// ============================================================
// Code_patch.js - 기존 Code.js 맨 끝에 붙여넣기
// 추가 기능: 견적 복사, 대시보드 통계, 동시편집 충돌 감지
// ============================================================

/**
 * 견적 복사(복제) - 기존 견적의 고객정보 + 아이템을 새 견적으로 복사
 * @param {string} sourceQuoteId - 복사할 원본 견적 ID
 * @param {string} adminPassword - 관리자 비밀번호
 * @returns {{ ok: boolean, quoteId: string, item_count: number }}
 */
function duplicateQuote(sourceQuoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var srcId = String(sourceQuoteId || "").trim();
  if (!srcId) throw new Error("sourceQuoteId required");

  var srcQuote = findQuoteRow_(srcId);
  if (!srcQuote) throw new Error("Source quote not found");

  // 새 견적 생성
  var newResult = createQuote(adminPassword);
  var newId = newResult.quoteId;

  // 고객/현장 정보 복사 (상태/토큰/날짜 등은 복사하지 않음)
  var copyFields = [
    "customer_name", "site_name", "contact_name", "contact_phone",
    "memo", "vat_included",
    "include_service_in_total", "include_option_in_total", "include_included_in_total"
  ];
  var patch = {};
  for (var i = 0; i < copyFields.length; i++) {
    var f = copyFields[i];
    if (srcQuote[f] !== undefined && srcQuote[f] !== null && String(srcQuote[f]).trim() !== "") {
      patch[f] = srcQuote[f];
    }
  }
  if (Object.keys(patch).length) {
    updateQuote_(newId, patch);
  }

  // 아이템 복사
  var srcItems = getItemsFromCache_(srcId);
  if (!srcItems) {
    srcItems = findItems_(srcId);
  }
  srcItems = normalizeItems_(srcItems || []);

  if (srcItems.length) {
    // 아이템 ID를 새로 발급하고 quote_id를 교체
    var newItems = [];
    for (var j = 0; j < srcItems.length; j++) {
      var item = cloneObj_(srcItems[j]);
      item.quote_id = newId;
      item.item_id = uuid_();
      newItems.push(item);
    }
    replaceItems_(newId, newItems);
    upsertQuoteItemsCache_(newId, newItems);
  }

  // 합계 재계산
  var s = getSettings_();
  var vatRate = Number(s.vat_rate || 0.1);
  var totals = calculateTotalsFromItems_(srcItems, patch.vat_included || "N", vatRate, {
    include_service_in_total: patch.include_service_in_total || "N",
    include_option_in_total: patch.include_option_in_total || "N",
    include_included_in_total: patch.include_included_in_total || "N"
  });
  updateQuote_(newId, {
    subtotal: totals.subtotal,
    vat: totals.vat,
    total: totals.total
  });

  bumpQuoteListVersion_();

  return {
    ok: true,
    quoteId: newId,
    source_quote_id: srcId,
    item_count: srcItems.length,
    customer_name: String(srcQuote.customer_name || ""),
    site_name: String(srcQuote.site_name || "")
  };
}


/**
 * 대시보드 통계 - 월별/상태별 요약 + 전환율
 * @param {string} adminPassword
 * @param {object} options - { months: number (기본 6) }
 * @returns {{ ok, monthly_stats, status_summary, conversion }}
 */
function getDashboardStats(adminPassword, options) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var opts = options || {};
  var monthsBack = Math.min(Math.max(Number(opts.months || 6), 1), 24);

  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      ok: true,
      monthly_stats: [],
      status_summary: { draft: 0, sent: 0, approved: 0, expired: 0, total: 0 },
      conversion: { sent_to_approved_pct: 0, draft_to_sent_pct: 0 },
      avg_quote_amount: 0,
      top_month_amount: 0
    };
  }

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var now = new Date();
  var cutoffDate = new Date(now.getFullYear(), now.getMonth() - monthsBack + 1, 1);

  var statusCount = { draft: 0, sent: 0, approved: 0, expired: 0 };
  var monthlyMap = {};
  var totalAmount = 0;
  var totalQuotes = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = String(row[cols.status] || "draft").toLowerCase().trim();
    var total = Number(row[cols.total] || 0);
    var createdAt = String(row[cols.created_at] || "");

    // 상태 집계
    if (status === "draft") statusCount.draft++;
    else if (status === "sent") statusCount.sent++;
    else if (status === "approved") statusCount.approved++;
    else statusCount.expired++;

    totalQuotes++;
    totalAmount += total;

    // 월별 집계
    var created = new Date(createdAt);
    if (isNaN(created.getTime())) continue;
    if (created < cutoffDate) continue;

    var monthKey = created.getFullYear() + "-" + String(created.getMonth() + 1).padStart(2, "0");
    if (!monthlyMap[monthKey]) {
      monthlyMap[monthKey] = { month: monthKey, count: 0, total: 0, approved_count: 0, approved_total: 0 };
    }
    monthlyMap[monthKey].count++;
    monthlyMap[monthKey].total += total;
    if (status === "approved") {
      monthlyMap[monthKey].approved_count++;
      monthlyMap[monthKey].approved_total += total;
    }
  }

  // 월별 배열로 변환 & 정렬
  var monthlyStats = [];
  for (var mk in monthlyMap) {
    if (!Object.prototype.hasOwnProperty.call(monthlyMap, mk)) continue;
    monthlyStats.push(monthlyMap[mk]);
  }
  monthlyStats.sort(function(a, b) { return a.month < b.month ? -1 : a.month > b.month ? 1 : 0; });

  // 전환율 계산
  var totalSentOrApproved = statusCount.sent + statusCount.approved;
  var sentToApproved = totalSentOrApproved > 0
    ? Math.round((statusCount.approved / totalSentOrApproved) * 100)
    : 0;
  var draftToSent = totalQuotes > 0
    ? Math.round((totalSentOrApproved / totalQuotes) * 100)
    : 0;

  // 최고 월 매출
  var topMonthAmount = 0;
  for (var m = 0; m < monthlyStats.length; m++) {
    if (monthlyStats[m].total > topMonthAmount) topMonthAmount = monthlyStats[m].total;
  }

  return {
    ok: true,
    monthly_stats: monthlyStats,
    status_summary: {
      draft: statusCount.draft,
      sent: statusCount.sent,
      approved: statusCount.approved,
      expired: statusCount.expired,
      total: totalQuotes
    },
    conversion: {
      sent_to_approved_pct: sentToApproved,
      draft_to_sent_pct: draftToSent
    },
    avg_quote_amount: totalQuotes > 0 ? Math.round(totalAmount / totalQuotes) : 0,
    top_month_amount: topMonthAmount
  };
}


/**
 * 동시편집 충돌 감지를 위한 saveQuote 래퍼
 * expectedUpdatedAt 파라미터를 받아, 서버측 updated_at과 비교
 * 불일치 시 에러를 발생시켜 클라이언트가 새로고침/병합하도록 유도
 *
 * @param {object} payload - 기존 saveQuote payload + payload.expectedUpdatedAt
 * @param {string} adminPassword
 * @returns {{ ok: boolean }}
 */
function saveQuoteWithConflictCheck(payload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var quoteId = String(payload && payload.quoteId || "").trim();
  if (!quoteId) throw new Error("quoteId required");

  var expectedUpdatedAt = String(payload && payload.expectedUpdatedAt || "").trim();

  if (expectedUpdatedAt) {
    var current = findQuoteRow_(quoteId);
    if (!current) throw new Error("Quote not found");

    var serverUpdatedAt = String(current.updated_at || "").trim();
    if (serverUpdatedAt && expectedUpdatedAt && serverUpdatedAt !== expectedUpdatedAt) {
      throw new Error(
        "CONFLICT: 다른 곳에서 이 견적이 수정되었어요. " +
        "저장 전에 새로고침하여 최신 내용을 확인해 주세요. " +
        "(서버: " + serverUpdatedAt + ", 기대: " + expectedUpdatedAt + ")"
      );
    }
  }

  return saveQuote(payload, adminPassword);
}


/**
 * 견적 삭제 (소프트 삭제 - status를 deleted로 변경)
 * @param {string} quoteId
 * @param {string} adminPassword
 * @returns {{ ok: boolean }}
 */
function archiveQuote(quoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  if (!qid) throw new Error("quoteId required");

  var existing = findQuoteRow_(qid);
  if (!existing) throw new Error("Quote not found");

  updateQuote_(qid, {
    status: "archived",
    updated_at: nowIso_()
  });
  bumpQuoteListVersion_();

  return { ok: true, quoteId: qid };
}


/**
 * 공정별 매출 비중 분석
 * @param {string} adminPassword
 * @param {object} options - { status_filter: "approved" | "all", limit: number }
 * @returns {{ ok, groups: [{ group_label, total, percentage, item_count }] }}
 */
function getGroupBreakdown(adminPassword, options) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var opts = options || {};
  var statusFilter = String(opts.status_filter || "all").toLowerCase();
  var limit = Math.min(Math.max(Number(opts.limit || 20), 1), 100);

  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, groups: [] };

  // 대상 견적 ID 수집
  var quoteIds = {};
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (var i = 0; i < values.length; i++) {
    var qid = String(values[i][cols.quote_id] || "").trim();
    var st = String(values[i][cols.status] || "").toLowerCase().trim();
    if (!qid) continue;
    if (statusFilter !== "all" && st !== statusFilter) continue;
    quoteIds[qid] = true;
  }

  // Items 시트에서 공정별 집계
  var shItems = getSheet_("Items");
  var hItems = getHeaders_("Items");
  var cItems = getColMap_("Items");
  var lastItemRow = shItems.getLastRow();
  if (lastItemRow < 2) return { ok: true, groups: [] };

  var itemValues = shItems.getRange(2, 1, lastItemRow - 1, hItems.length).getValues();
  var groupMap = {};
  var grandTotal = 0;

  for (var j = 0; j < itemValues.length; j++) {
    var row = itemValues[j];
    var itemQuoteId = String(row[cItems.quote_id] || "").trim();
    if (!quoteIds[itemQuoteId]) continue;

    var groupLabel = String(row[cItems.group_label] || row[cItems.process] || "기타").trim();
    var price = Number(row[cItems.price] || row[cItems.amount] || 0);
    if (isNaN(price)) price = 0;

    if (!groupMap[groupLabel]) {
      groupMap[groupLabel] = { group_label: groupLabel, total: 0, item_count: 0 };
    }
    groupMap[groupLabel].total += price;
    groupMap[groupLabel].item_count++;
    grandTotal += price;
  }

  var groups = [];
  for (var gl in groupMap) {
    if (!Object.prototype.hasOwnProperty.call(groupMap, gl)) continue;
    var g = groupMap[gl];
    g.percentage = grandTotal > 0 ? Math.round((g.total / grandTotal) * 1000) / 10 : 0;
    groups.push(g);
  }
  groups.sort(function(a, b) { return b.total - a.total; });
  if (groups.length > limit) groups = groups.slice(0, limit);

  return { ok: true, groups: groups, grand_total: grandTotal };
}


/**
 * 견적 내보내기 (CSV 형태)
 * @param {string} quoteId
 * @param {string} adminPassword
 * @returns {{ ok, csv_content, filename }}
 */
function exportQuoteCsv(quoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  if (!qid) throw new Error("quoteId required");

  var q = findQuoteRow_(qid);
  if (!q) throw new Error("Quote not found");

  var items = getItemsFromCache_(qid);
  if (!items) items = findItems_(qid);
  items = normalizeItems_(items || []);

  var csvHeaders = ["공정", "품명", "세부내역", "규격", "수량", "단위", "단가", "금액", "가격유형", "비고"];
  var rows = [csvHeaders.join(",")];

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var row = [
      csvEscape_(it.group_label || it.process || ""),
      csvEscape_(it.name || it.material || ""),
      csvEscape_(it.location || ""),
      csvEscape_(it.detail || it.spec || ""),
      String(it.qty || 0),
      csvEscape_(it.unit || ""),
      String(it.unit_price || 0),
      String(it.price || it.amount || 0),
      csvEscape_(it.price_type || "NORMAL"),
      csvEscape_(it.note || "")
    ];
    rows.push(row.join(","));
  }

  var customerName = String(q.customer_name || "견적").trim().replace(/[^가-힣a-zA-Z0-9_-]/g, "_");
  var filename = customerName + "_견적_" + qid.slice(0, 8) + ".csv";

  return {
    ok: true,
    csv_content: rows.join("\n"),
    filename: filename,
    item_count: items.length
  };
}

function csvEscape_(val) {
  var s = String(val || "");
  if (s.indexOf(",") >= 0 || s.indexOf('"') >= 0 || s.indexOf("\n") >= 0) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}


/**
 * 최근 활동 요약 (대시보드 피드용)
 * @param {string} adminPassword
 * @param {number} limit - 최대 항목 수 (기본 20)
 * @returns {{ ok, activities: [{ type, quote_id, customer_name, message, timestamp }] }}
 */
function getRecentActivities(adminPassword, limit) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var maxItems = Math.min(Math.max(Number(limit || 20), 1), 100);
  var activities = [];

  // 최근 견적
  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    var startRow = Math.max(2, lastRow - 49);
    var values = sh.getRange(startRow, 1, lastRow - startRow + 1, headers.length).getValues();

    for (var i = values.length - 1; i >= 0 && activities.length < maxItems; i--) {
      var row = values[i];
      var qid = String(row[cols.quote_id] || "").trim();
      var name = String(row[cols.customer_name] || "").trim();
      var status = String(row[cols.status] || "").toLowerCase();
      var updatedAt = String(row[cols.updated_at] || row[cols.created_at] || "");

      if (!qid) continue;

      var message = "";
      if (status === "approved") {
        message = (name || "고객") + " 견적이 확정되었어요.";
      } else if (status === "sent") {
        message = (name || "고객") + " 견적이 발송되었어요.";
      } else {
        message = (name || "고객") + " 견적이 수정되었어요.";
      }

      activities.push({
        type: "quote_" + status,
        quote_id: qid,
        customer_name: name,
        message: message,
        timestamp: updatedAt
      });
    }
  }

  // 최근 고객 메모
  try {
    var shNotes = getSheet_("CustomerNotes");
    var hNotes = getHeaders_("CustomerNotes");
    var cNotes = getColMap_("CustomerNotes");
    var lastNoteRow = shNotes.getLastRow();

    if (lastNoteRow >= 2) {
      var noteStart = Math.max(2, lastNoteRow - 19);
      var noteValues = shNotes.getRange(noteStart, 1, lastNoteRow - noteStart + 1, hNotes.length).getValues();

      for (var n = noteValues.length - 1; n >= 0 && activities.length < maxItems * 2; n--) {
        var noteRow = noteValues[n];
        var noteQuoteId = String(noteRow[cNotes.quote_id] || "").trim();
        var noteText = String(noteRow[cNotes.note] || "").trim();
        var noteAt = String(noteRow[cNotes.created_at] || "");

        if (!noteQuoteId || !noteText) continue;

        activities.push({
          type: "customer_note",
          quote_id: noteQuoteId,
          customer_name: "",
          message: "고객 메모: " + noteText.substring(0, 60) + (noteText.length > 60 ? "..." : ""),
          timestamp: noteAt
        });
      }
    }
  } catch (e) {}

  // 시간순 정렬 (최신순)
  activities.sort(function(a, b) {
    return (b.timestamp || "") > (a.timestamp || "") ? 1 : -1;
  });

  return {
    ok: true,
    activities: activities.slice(0, maxItems)
  };
}
