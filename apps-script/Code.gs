// [TEST부스] Code.gs 성적표생성DB Set_ID WM형식 정상화 16350-16419
/* =========================================================
   TEST / REAL 환경설정 (이 블록만 수정)
   - 리얼부스 운영 기본값: REAL
   - TEST 전환 시 WM_ENV 값만 'TEST'로 변경
   - 하단 기존 로직은 WM_LMS_SPREADSHEET_ID / WM_CONTENT_SPREADSHEET_ID 공통 변수만 사용
========================================================= */
const WM_ENV = 'REAL';  // 'REAL' 또는 'TEST'

const WM_ENV_CONFIG = {
  REAL: {
    LMS_SPREADSHEET_ID: '1BMEP17xhnwW1fq2pfZlfNCJ3t9z7UXQ_BF3XxxmNOos',
    CONTENT_SPREADSHEET_ID: '1grMQjmNR6JMm7YFyMr8HPj_ov3DJkXU6vHA_aAWG3fk'
  },
  TEST: {
    LMS_SPREADSHEET_ID: '1EgJbYnP2Yc9QyooTVZLGWein-VhqmXUiFZVHj8PbYPE',
    CONTENT_SPREADSHEET_ID: '1-lddlzcMgv5OqFXucuqHSbrRIdQtVBIrSM2ZLV7mXN8'
  }
};

var WM_LMS_SPREADSHEET_ID = WM_ENV_CONFIG[WM_ENV].LMS_SPREADSHEET_ID;
var WM_CONTENT_SPREADSHEET_ID = WM_ENV_CONFIG[WM_ENV].CONTENT_SPREADSHEET_ID;

/* =========================================================
   아래부터 기존 Code.gs 로직
========================================================= */

/* WM_CODEGS_STABLE_LEARNING_FLOW_260612_V1 */
/*
WM_CODEGS_ROUTER_INTEGRATED_V3
적용 위치: Code.gs 전체 교체용
역할: doGet/doPost 단일 관제센터 + 기존 map/study/dev/studentLogin/getSetData/saveLearningRecord/health 유지 + integrated 진입 추가
주의: Apps Script 프로젝트 안에서 doGet(e)는 Code.gs 하나만 사용합니다. Integrated_Code.gs에는 doGet을 두지 않습니다.
*/

/*
WM_LMS_OFFICIAL_RECORD_RULES_20260608_V1
[확정 공식]
1. 점수 = 한영주관식 점수. 평균점수는 사용하지 않습니다.
2. 총소요시간 = 완료된 Step 시간 합산. Step별 시간은 각 Step 단독시간입니다.
3. Test_총시간 = 4개 시험 시간만 합산합니다.
4. 세트진행률 = 현재 학습 세트의 Step 진행 상태를 표시합니다.
5. 레벨진행률 = 완료세트 ÷ 전체세트 × 100, 소수점 반올림.
6. 학습기록ID = R + YY + MM + 3자리 순번.
7. Class 표시는 반ID가 아니라 실제 반명을 우선합니다.
*/

/* WM_CODEGS_MAP_STUDENTNAME_ROOT_DATA_FIX_V13 */
/* WM_GET_LEARNING_MAP_FULL_ROUTE_AND_RUN_FIX_V1: getLearningMap JSON/API + google.script.run + integrated direct separated route. */
/* WM_CODEGS_LMS_SAVE_API_V1: Study 완료 기록을 2.학습기록_DB에 저장. */
/* WM_STEP_RESUME_PROGRESS_V1: Step 완료 기준 저장 + 재진입 시 다음 Step 조회 API 추가. */
/* WM_LMS_SPREADSHEET_RUNTIME_CACHE_V1
 * 한 API 요청에서 DB마다 같은 스프레드시트를 반복 openById 하던 지연을 제거합니다.
 * 실행 컨테이너 안에서는 한 번 연 객체를 재사용하고, 재사용 실패 시에만 다시 엽니다. */
var WM_LMS_SPREADSHEET_RUNTIME_CACHE_ = null;
function getLmsSpreadsheet_() {
  if (WM_LMS_SPREADSHEET_RUNTIME_CACHE_) {
    try {
      WM_LMS_SPREADSHEET_RUNTIME_CACHE_.getId();
      return WM_LMS_SPREADSHEET_RUNTIME_CACHE_;
    } catch (cacheError) {
      WM_LMS_SPREADSHEET_RUNTIME_CACHE_ = null;
    }
  }
  WM_LMS_SPREADSHEET_RUNTIME_CACHE_ = SpreadsheetApp.openById(WM_LMS_SPREADSHEET_ID);
  return WM_LMS_SPREADSHEET_RUNTIME_CACHE_;
}

/* WM_SPEED_CACHE_LAYER_20260703_V1
 * 속도개선 전용 캐시 레이어입니다.
 * 원본 DB 구조/저장 공식은 변경하지 않고, 반복 조회만 줄입니다.
 */
function wmGetScriptCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (err) {
    return null;
  }
}

function wmCacheGetJson_(key) {
  var cache = wmGetScriptCache_();
  if (!cache || !key) return null;
  try {
    var text = cache.get(String(key));
    return text ? JSON.parse(text) : null;
  } catch (err) {
    return null;
  }
}

function wmCachePutJson_(key, value, ttlSeconds) {
  var cache = wmGetScriptCache_();
  if (!cache || !key) return;
  try {
    var text = JSON.stringify(value);
    if (text.length > 90000) return;
    cache.put(String(key), text, Number(ttlSeconds || 300));
  } catch (err) {}
}

function wmCacheRemove_(key) {
  var cache = wmGetScriptCache_();
  if (!cache || !key) return;
  try { cache.remove(String(key)); } catch (err) {}
}

function wmCacheKey_(name, id) {
  return ['WM', WM_ENV, name, String(id || '').trim().toUpperCase()].join(':');
}

function wmClearRuntimeCachesForStudent_(studentId) {
  /* WM_CURRENT_PROGRESS_CACHE_CLEAR_BATCH_V1_20260821
   * 완료 직후 CacheService.remove를 여러 번 왕복하지 않고 한 번에 제거합니다.
   * 현재진행_DB 저장 완료 뒤 응답 지연만 줄이며 저장 공식은 변경하지 않습니다. */
  var sid = String(studentId || '').trim().toUpperCase();
  var cache = wmGetScriptCache_();
  if (!cache || !sid) return;
  var keys = [
    wmCacheKey_('STUDENT_PROFILE', sid),
    wmCacheKey_('STUDENT_MODE', sid),
    wmCacheKey_('CP_SET_MAP', 'ALL'),
    wmCacheKey_('CURRENT_PROGRESS_STUDENT', sid),
    wmCacheKey_('MAP_RECORDS', sid),
    wmCacheKey_('LMS_CUMULATIVE_WRONG', sid)
  ];
  try { cache.removeAll(keys); } catch (err) {
    keys.forEach(function(key){ wmCacheRemove_(key); });
  }
}

/* WM_CLASS_SAVE_BATCH_CACHE_CLEAR_V1
 * 반등록 시 학생마다 CacheService.remove를 반복하지 않고 한 번에 제거합니다. */
function wmClearRuntimeCachesForStudents_(studentIds) {
  var cache = wmGetScriptCache_();
  if (!cache) return;
  var keys = [wmCacheKey_('CP_SET_MAP', 'ALL')];
  (studentIds || []).forEach(function(studentId){
    var sid = String(studentId || '').trim().toUpperCase();
    if (!sid) return;
    keys.push(wmCacheKey_('STUDENT_PROFILE', sid));
    keys.push(wmCacheKey_('STUDENT_MODE', sid));
    keys.push(wmCacheKey_('CURRENT_PROGRESS_STUDENT', sid));
    keys.push(wmCacheKey_('MAP_RECORDS', sid));
  });
  try { cache.removeAll(keys); } catch (err) {
    keys.forEach(function(key){ try { cache.remove(key); } catch (ignore) {} });
  }
}

/* WM_MAP_SINGLE_STUDENT_ROWS_SPEED_V1_20260713
 * Map 조회에서 대형 DB 전체 열을 읽지 않고 학생ID 열 1개만 확인한 뒤,
 * 해당 학생 행이 이어진 구간만 묶어서 읽습니다.
 */
function wmGetStudentRowNumbersFast_(sheet, headers, studentId) {
  if (!sheet || !headers || !studentId) return [];
  var idxStudent = headers.indexOf('학생ID');
  var lastRow = sheet.getLastRow();
  if (idxStudent < 0 || lastRow < 2) return [];

  var wanted = String(studentId || '').trim().toUpperCase();
  var idValues = sheet.getRange(2, idxStudent + 1, lastRow - 1, 1).getDisplayValues();
  var rowNumbers = [];
  for (var i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0] || '').trim().toUpperCase() === wanted) rowNumbers.push(i + 2);
  }
  return rowNumbers;
}

function wmReadGroupedRowsFast_(sheet, rowNumbers, lastCol) {
  if (!sheet || !rowNumbers || !rowNumbers.length || lastCol < 1) return [];
  var groupCount = 1;
  for (var g = 1; g < rowNumbers.length; g++) {
    if (rowNumbers[g] !== rowNumbers[g - 1] + 1) groupCount += 1;
  }
  /* 학생행이 여러 곳에 흩어진 경우 다수의 시트 호출이 오히려 느려지므로
   * 12구간 초과 시 기존 1회 전체조회 방식으로 안전하게 전환합니다. */
  if (groupCount > 12) {
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  }
  var result = [];
  var start = rowNumbers[0];
  var end = start;

  function readBlock_() {
    var block = sheet.getRange(start, 1, end - start + 1, lastCol).getDisplayValues();
    for (var b = 0; b < block.length; b++) result.push(block[b]);
  }

  for (var i = 1; i < rowNumbers.length; i++) {
    if (rowNumbers[i] === end + 1) {
      end = rowNumbers[i];
    } else {
      readBlock_();
      start = rowNumbers[i];
      end = start;
    }
  }
  readBlock_();
  return result;
}


/* =========================================================
   🔒 WM SPEED LOCK V1 - SERVER ROUTE BINDING / 실제 실행 연결
   로그인·학습맵 초기표시·학습기록 동시표시 서버 함수를
   동결된 Registry를 통해 doGet 라우터가 직접 실행합니다.
   ========================================================= */
var WM_SPEED_LOCK_SERVER_V1 = Object.freeze({
  version: 'WM_LOGIN_SPEED_LOCK_V2.0',
  date: '2026-07-13',
  login: studentLogin,
  loginOptimization: Object.freeze({
    singleRowLookup: true,
    rowCache: true,
    headerCache: true,
    sessionBatchWrite: true
  }),
  mapCache: getLearningMapCacheData,
  mapRecords: getLearningRecordsForMapData
});

function wmGetSpeedLockServerStatus_() {
  var registry = WM_SPEED_LOCK_SERVER_V1;
  var result = {
    version: registry.version,
    date: registry.date,
    login: typeof registry.login === 'function',
    loginOptimization: !!registry.loginOptimization &&
      registry.loginOptimization.singleRowLookup === true &&
      registry.loginOptimization.rowCache === true &&
      registry.loginOptimization.headerCache === true &&
      registry.loginOptimization.sessionBatchWrite === true &&
      Object.isFrozen(registry.loginOptimization),
    mapCache: typeof registry.mapCache === 'function',
    mapRecords: typeof registry.mapRecords === 'function'
  };
  result.loginVerified = result.login && result.loginOptimization && Object.isFrozen(registry);
  result.verified = result.loginVerified && result.mapCache && result.mapRecords;
  return result;
}

/* =========================================================
   🔒 WM STUDY SPEED SERVER LOCK V1 / 실제 실행 연결
   세트조회·빠른 복귀조회·최종저장·현재진행 갱신 경로를 동결합니다.
   ========================================================= */
var WM_STUDY_SPEED_SERVER_LOCK_V1 = Object.freeze({
  version: 'WM_STUDY_SPEED_LOCK_V1.0',
  date: '2026-07-11',
  setDataFast: getSetData,
  progressFast: getStudyProgressFast,
  finalSaveFast: wmSaveLearningRecordOfficialTotalV2_,
  currentProgressFast: wmUpsertCurrentProgress_
});

function wmGetStudySpeedServerLockStatus_() {
  var registry = WM_STUDY_SPEED_SERVER_LOCK_V1;
  var result = {
    version: registry.version,
    date: registry.date,
    setDataFast: typeof registry.setDataFast === 'function',
    progressFast: typeof registry.progressFast === 'function',
    finalSaveFast: typeof registry.finalSaveFast === 'function',
    currentProgressFast: typeof registry.currentProgressFast === 'function'
  };
  result.verified = result.setDataFast && result.progressFast && result.finalSaveFast &&
    result.currentProgressFast && Object.isFrozen(registry);
  return result;
}

/* =========================================================
   🔒 WM SETTING CENTER SERVER LOCK V1 / 실제 실행 연결
   설정센터 조회·중복확인·교사등록·상태저장·반등록 API를
   동결 Registry를 통해 doGet 라우터가 직접 실행합니다.
   ========================================================= */
var WM_SETTING_CENTER_SERVER_LOCK_V1 = Object.freeze({
  version: 'WM_SETTING_CENTER_LOCK_V1.0',
  date: '2026-07-10',
  getSettingCenter: wmGetSettingCenterForLms_,
  checkTeacherId: wmCheckSettingTeacherIdForLms_,
  saveTeacher: wmSaveSettingTeacherForLms_,
  saveTeacherStatus: wmSaveSettingTeacherStatusForLms_,
  saveClass: wmSaveSettingClassForLms_
});

function wmGetSettingCenterServerLockStatus_() {
  var registry = WM_SETTING_CENTER_SERVER_LOCK_V1;
  var result = {
    version: registry.version,
    date: registry.date,
    getSettingCenter: typeof registry.getSettingCenter === 'function',
    checkTeacherId: typeof registry.checkTeacherId === 'function',
    saveTeacher: typeof registry.saveTeacher === 'function',
    saveTeacherStatus: typeof registry.saveTeacherStatus === 'function',
    saveClass: typeof registry.saveClass === 'function'
  };
  result.verified =
    result.getSettingCenter &&
    result.checkTeacherId &&
    result.saveTeacher &&
    result.saveTeacherStatus &&
    result.saveClass &&
    Object.isFrozen(registry);
  return result;
}

/* =========================================================
   🔒 WM CODE.GS STUDENT ASSIGNMENT SERVER LOCK V1 / 실제 실행 연결
   LMS 학생관리의 최초배정·현재세트·다음레벨완료조건선택·
   저장·현재세트 기준 다음레벨 +1 계산을 동결 Registry로 연결합니다.
   ========================================================= */
var WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1 = Object.freeze({
  version: 'WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1.0',
  date: '2026-07-14',
  FIRST_ASSIGNMENT: Object.freeze({
    save: saveStudentLearningModeApi_
  }),
  CURRENT_SET: Object.freeze({
    read: buildCurrentProgressSetMapForLms_
  }),
  NEXT_LEVEL_COMPLETE_CONDITION: Object.freeze({
    normalize: normalizeLevelCompleteCondition_
  }),
  SAVE: Object.freeze({
    studentAssignment: saveStudentLearningModeApi_
  }),
  NEXT_LEVEL_PLUS_ONE: Object.freeze({
    calculate: wmNextLevelFromCurrentSetForLms_
  })
});

function wmGetCodeGsStudentAssignmentLockStatus_() {
  var registry = WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1;
  var moduleNames = ['FIRST_ASSIGNMENT','CURRENT_SET','NEXT_LEVEL_COMPLETE_CONDITION','SAVE','NEXT_LEVEL_PLUS_ONE'];
  var modules = {};
  var allVerified = Object.isFrozen(registry);

  moduleNames.forEach(function(moduleName) {
    var group = registry[moduleName] || {};
    var functions = {};
    var groupVerified = Object.isFrozen(group);
    Object.keys(group).forEach(function(functionName) {
      functions[functionName] = typeof group[functionName] === 'function';
      groupVerified = groupVerified && functions[functionName];
    });
    modules[moduleName] = { verified: groupVerified, functions: functions };
    allVerified = allVerified && groupVerified;
  });

  allVerified = allVerified &&
    registry.FIRST_ASSIGNMENT.save === saveStudentLearningModeApi_ &&
    registry.CURRENT_SET.read === buildCurrentProgressSetMapForLms_ &&
    registry.NEXT_LEVEL_COMPLETE_CONDITION.normalize === normalizeLevelCompleteCondition_ &&
    registry.SAVE.studentAssignment === saveStudentLearningModeApi_ &&
    registry.NEXT_LEVEL_PLUS_ONE.calculate === wmNextLevelFromCurrentSetForLms_;

  return {
    version: registry.version,
    date: registry.date,
    modules: modules,
    registryFrozen: Object.isFrozen(registry),
    allVerified: allVerified
  };
}

/* WM CODE.GS LEVEL COMPLETE CONDITION SERVER LOCK V1 */
var WM_CODEGS_LEVEL_COMPLETE_CONDITION_LOCK_V1 = Object.freeze({
  version:'WM_CODEGS_LEVEL_COMPLETE_CONDITION_LOCK_V1.0',
  CURRENT_PROGRESS:Object.freeze({read:buildCurrentProgressSetMapForLms_}),
  CONDITION:Object.freeze({normalize:normalizeLevelCompleteCondition_}),
  SAVE:Object.freeze({studentLearningMode:saveStudentLearningModeApi_})
});

function wmGetCodeGsLevelCompleteConditionLockStatus_(){
  var registry = WM_CODEGS_LEVEL_COMPLETE_CONDITION_LOCK_V1;
  var moduleNames = ['CURRENT_PROGRESS','CONDITION','SAVE'];
  var modules = {};
  var allVerified = Object.isFrozen(registry);
  moduleNames.forEach(function(moduleName){
    var group = registry[moduleName] || {};
    var functions = {};
    var groupVerified = Object.isFrozen(group);
    Object.keys(group).forEach(function(functionName){
      functions[functionName] = typeof group[functionName] === 'function';
      groupVerified = groupVerified && functions[functionName];
    });
    modules[moduleName] = {verified:groupVerified,functions:functions};
    allVerified = allVerified && groupVerified;
  });
  allVerified = allVerified
    && registry.CURRENT_PROGRESS.read === buildCurrentProgressSetMapForLms_
    && registry.CONDITION.normalize === normalizeLevelCompleteCondition_
    && registry.SAVE.studentLearningMode === saveStudentLearningModeApi_;
  return {version:registry.version,modules:modules,registryFrozen:Object.isFrozen(registry),allVerified:allVerified};
}

/* WM CODE.GS LMS CORE ACTUAL FUNCTION LOCK V1 */
const WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1 = Object.freeze({
  version: 'WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.0',
  SETTING_CENTER: Object.freeze({
    getSettingCenter: wmGetSettingCenterForLms_,
    checkTeacherId: wmCheckSettingTeacherIdForLms_,
    saveTeacher: wmSaveSettingTeacherForLms_,
    saveTeacherStatus: wmSaveSettingTeacherStatusForLms_,
    saveClass: wmSaveSettingClassForLms_
  }),
  STUDENT_DB: Object.freeze({
    load: getStudentsForLmsApi_,
    save: saveStudentLearningModeApi_,
    currentSet: buildCurrentProgressSetMapForLms_,
    actorFilter: wmFilterStudentsForActor_
  }),
  STUDY_RECORD_DB: Object.freeze({
    load: getLearningRecordsForLmsApi_,
    save: saveLearningRecord,
    readRows: wmReadLmsSheetAsObjects_,
    actorFilter: wmFilterRowsByStudentSet_
  })
});

function wmGetCodeGsLmsCoreActualLockStatus_() {
  var registry = WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1;
  var moduleNames = ['SETTING_CENTER','STUDENT_DB','STUDY_RECORD_DB'];
  var modules = {};
  var allVerified = Object.isFrozen(registry);
  moduleNames.forEach(function(moduleName) {
    var group = registry[moduleName] || {};
    var functions = {};
    var verified = Object.isFrozen(group);
    Object.keys(group).forEach(function(key) {
      functions[key] = typeof group[key] === 'function';
      verified = verified && functions[key];
    });
    modules[moduleName] = {verified:verified, functions:functions};
    allVerified = allVerified && verified;
  });
  allVerified = allVerified
    && registry.SETTING_CENTER.getSettingCenter === wmGetSettingCenterForLms_
    && registry.SETTING_CENTER.checkTeacherId === wmCheckSettingTeacherIdForLms_
    && registry.SETTING_CENTER.saveTeacher === wmSaveSettingTeacherForLms_
    && registry.SETTING_CENTER.saveTeacherStatus === wmSaveSettingTeacherStatusForLms_
    && registry.SETTING_CENTER.saveClass === wmSaveSettingClassForLms_
    && registry.STUDENT_DB.load === getStudentsForLmsApi_
    && registry.STUDENT_DB.save === saveStudentLearningModeApi_
    && registry.STUDENT_DB.currentSet === buildCurrentProgressSetMapForLms_
    && registry.STUDENT_DB.actorFilter === wmFilterStudentsForActor_
    && registry.STUDY_RECORD_DB.load === getLearningRecordsForLmsApi_
    && registry.STUDY_RECORD_DB.save === saveLearningRecord
    && registry.STUDY_RECORD_DB.readRows === wmReadLmsSheetAsObjects_
    && registry.STUDY_RECORD_DB.actorFilter === wmFilterRowsByStudentSet_;
  return {version:registry.version, modules:modules, registryFrozen:Object.isFrozen(registry), allVerified:allVerified};
}

function wmIsCodeGsLmsCoreProtectedRequest_(action, mode) {
  var key = String(action || mode || '').trim();
  return [
    'saveLearningRecord','saveStudentLearningMode','students','records','learningRecords','gradeLevels','gradeLevel',
    'getSettingCenter','checkTeacherId','saveSettingTeacher','saveSettingTeacherStatus','saveSettingClass'
  ].indexOf(key) !== -1;
}

/* =========================================================
   🔒 WM LEVEL TRANSITION NOTICE SERVER LOCK V1 / 실제 실행 연결
   TEST 완료 저장 후 회차계속·레벨완료·13레벨 전체완료 판정을
   동결 Registry로 연결합니다.
   ========================================================= */
var WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1 = Object.freeze({
  version: 'WM_LEVEL_TRANSITION_NOTICE_LOCK_V1.0',
  date: '2026-07-14',
  TRANSITION: Object.freeze({
    evaluate: wmBuildOfficialSequentialTransitionAfterTest_,
    levelCompletionState: wmGetLockedLevelCompletionState_
  })
});

function wmGetLevelTransitionNoticeServerLockStatus_() {
  var registry = WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1;
  var actualLock = Object.isFrozen(registry) &&
    Object.isFrozen(registry.TRANSITION) &&
    registry.TRANSITION.evaluate === wmBuildOfficialSequentialTransitionAfterTest_ &&
    registry.TRANSITION.levelCompletionState === wmGetLockedLevelCompletionState_;
  return {
    version: registry.version,
    date: registry.date,
    transitionEngine: actualLock,
    roundContinue: actualLock,
    levelComplete: actualLock,
    courseComplete: actualLock,
    registryFrozen: Object.isFrozen(registry),
    verified: actualLock
  };
}

/* =========================================================
   🔒 WM STUDY RUNTIME EXECUTION SERVER LOCK V1 / 실제 실행 연결
   Study에서 호출하는 세트조회·진행조회·저장·현재진행 갱신·
   회차/레벨 전환 판정의 실제 서버 함수 참조를 고정합니다.
   ========================================================= */
var WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1 = Object.freeze({
  'WM5-1-10|favorite':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158889/WM5-1-10_favorite_v2_q3zx3e.mp3',
  'WM5-2-7|club':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158405/WM5-2-7_club_v2_zvnxwx.mp3',
  'WM5-2-7|few':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158956/WM5-2-7_few_v2_o5u6uq.mp3',
  'WM5-2-7|love':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158994/WM5-2-7_love_v2_c6gb1f.mp3',
  'WM5-2-7|remember':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158997/WM5-2-7_remember_v2_syvpyn.mp3',
  'WM5-2-7|scissors':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159003/WM5-2-7_scissors_v2_ex9qjc.mp3',
  'WM5-2-7|everywhere':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784158927/WM5-2-7_everywhere_v2_rqi8ls.mp3',
  'WM6-2-3|post office':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159018/WM6-2-3_post-office_v2_f6ts1o.mp3',
  'WM6-2-8|earache':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159023/WM6-2-8_earache_v2_ribzwv.mp3',
  'WM7-1-9|blond':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159027/WM7-1-9_blond_v2_vdwvki.mp3',
  'WM7-2-1|curry':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159029/WM7-2-1_curry_v2_qaawed.mp3',
  'WM7-2-4|recycle':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159031/WM7-2-4_recycle_v2_v981or.mp3',
  'WM8-3-1|suspension':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159051/WM8-3-1__suspension_v2_zd6gqe.mp3',
  'WM9-1-3|plagiarism':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159056/WM9-1-3_plagiarism_v2_ip8yxh.mp3',
  'WM11-4-10|ridiculous':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159066/WM11-4-10_ridiculous_v2_i8sosi.mp3',
  'WM12-1-5|statistical':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159068/WM12-1-5_statistical_v2_igdd6o.mp3',
  'WM12-1-6|presumption':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159074/WM12-1-6_presumption_v2_c69p6j.mp3',
  'WM12-2-7|malnourished':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159076/WM12-2-7_malnourished_v2_icy5ag.mp3',
  'WM12-2-8|melancholic':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159098/WM12-2-8_melancholic_v2_rx1ne5.mp3',
  'WM12-2-8|controvert':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159078/WM12-2-8_controvert_v2_cytwbq.mp3',
  'WM13-1-6|uniqueness':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159101/WM13-1-6_uniqueness_v2_gkk8rt.mp3',
  'WM13-2-7|dither':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159106/WM13-2-7_dither_v2_yo0vqz.mp3',
  'WM13-3-3|penal':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159109/WM13-3-3_penal_v2_isliyr.mp3',
  'WM13-3-7|breed':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159111/WM13-3-7_breed_v2_xqgap1.mp3',
  'WM13-4-8|progenitor':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159119/WM13-4-8_progenitor_v2_g0fvnr.mp3',
  'WM13-4-8|application':'https://res.cloudinary.com/dxb7rmph3/video/upload/v1784159117/WM13-4-8_application_v2_neiryh.mp3'
});

function wmNormalizeWordAudioOverrideServerKey_(setId, word) {
  return String(setId || '').trim().toUpperCase() + '|' +
    String(word || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function wmGetWordAudioOverrideServerUrl_(setId, word) {
  return WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1[
    wmNormalizeWordAudioOverrideServerKey_(setId, word)
  ] || '';
}

function wmVerifyWordAudioOverrideServerLock_() {
  var keys = Object.keys(WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1);
  return Object.isFrozen(WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1) &&
    keys.length === 26 && keys.every(function(key) {
      return typeof WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1[key] === 'string' &&
        WM_WORD_AUDIO_OVERRIDE_SERVER_URLS_V1[key].trim() !== '';
    });
}

var WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1 = Object.freeze({
  version: 'WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.0',
  date: '2026-07-16',
  BOOT_CONTEXT: Object.freeze({
    studentInfo: getStudentBasicInfoForMap_,
    learningMode: getStudentLearningMode_
  }),
  LEARNING_MODE_API: Object.freeze({
    read: getStudentLearningModeApi_
  }),
  SET_DATA: Object.freeze({
    load: getSetData
  }),
  WORD_AUDIO_OVERRIDE: Object.freeze({
    normalize: wmNormalizeWordAudioOverrideServerKey_,
    resolve: wmGetWordAudioOverrideServerUrl_,
    verify: wmVerifyWordAudioOverrideServerLock_
  }),
  PROGRESS: Object.freeze({
    fast: getStudyProgressFast,
    legacy: getLearningProgress
  }),
  SAVE: Object.freeze({
    finalSave: wmSaveLearningRecordOfficialTotalV2_,
    checkpointFast: wmSaveStudyCheckpointFast_,
    fastBuildOrUpdate: wmFastBuildOrUpdateLearningRecord_,
    fastStudyTimeFields: wmFastStudyTimeFields_,
    fastMergeStudyTimes: wmFastMergeStudyTimesIntoRecord_,
    fastApplyFinalTestOnly: wmFastApplyFinalTestOnly_,
    officialTimeFields: wmBuildOfficialLearningTimeFieldsV2_,
    buildRecord: wmBuildLearningRecordOfficialTotalV2_,
    officialRound: wmApplyOfficialCompleteRoundToRecord_,
    findCurrentRecord: wmFindLatestIncompleteLearningRecordRowByStudentSet_,
    mergeCurrentRecord: wmMergeLatestIncompleteSetTimesForNewRecord_
  }),
  CURRENT_PROGRESS: Object.freeze({
    upsert: wmUpsertCurrentProgress_,
    fields: wmBuildCurrentProgressCacheFields_,
    syncStudentSet: wmSyncStudentCurrentSetAfterOfficialCompletion_,
    clearCache: wmClearRuntimeCachesForStudent_
  }),
  TRANSITION: Object.freeze({
    evaluate: wmBuildOfficialSequentialTransitionAfterTest_,
    levelCompletionState: wmGetLockedLevelCompletionState_
  })
});

function wmGetStudyRuntimeExecutionServerLockStatus_() {
  var registry = WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1;
  var moduleNames = ['BOOT_CONTEXT','LEARNING_MODE_API','SET_DATA','WORD_AUDIO_OVERRIDE','PROGRESS','SAVE','CURRENT_PROGRESS','TRANSITION'];
  var modules = {};
  var allVerified = Object.isFrozen(registry);

  moduleNames.forEach(function(moduleName) {
    var group = registry[moduleName] || {};
    var functions = {};
    var groupVerified = Object.isFrozen(group);
    Object.keys(group).forEach(function(functionName) {
      functions[functionName] = typeof group[functionName] === 'function';
      groupVerified = groupVerified && functions[functionName];
    });
    if (moduleName === 'WORD_AUDIO_OVERRIDE') {
      groupVerified = groupVerified && registry.WORD_AUDIO_OVERRIDE.verify();
      modules[moduleName] = {
        verified:groupVerified,
        functions:functions,
        entries:groupVerified ? '26/26' : 'CHECK REQUIRED'
      };
    } else {
      modules[moduleName] = {verified:groupVerified, functions:functions};
    }
    allVerified = allVerified && groupVerified;
  });

  allVerified = allVerified &&
    registry.BOOT_CONTEXT.studentInfo === getStudentBasicInfoForMap_ &&
    registry.BOOT_CONTEXT.learningMode === getStudentLearningMode_ &&
    registry.LEARNING_MODE_API.read === getStudentLearningModeApi_ &&
    registry.SET_DATA.load === WM_CODEGS_ALL_LOCK_V1.STUDY.setData &&
    registry.WORD_AUDIO_OVERRIDE.normalize === wmNormalizeWordAudioOverrideServerKey_ &&
    registry.WORD_AUDIO_OVERRIDE.resolve === wmGetWordAudioOverrideServerUrl_ &&
    registry.WORD_AUDIO_OVERRIDE.verify === wmVerifyWordAudioOverrideServerLock_ &&
    registry.WORD_AUDIO_OVERRIDE.verify() &&
    registry.PROGRESS.fast === WM_CODEGS_ALL_LOCK_V1.STUDY.progressFast &&
    registry.PROGRESS.legacy === WM_CODEGS_ALL_LOCK_V1.STUDY.progressLegacy &&
    registry.SAVE.finalSave === WM_CODEGS_ALL_LOCK_V1.STUDY.finalSave &&
    registry.SAVE.checkpointFast === wmSaveStudyCheckpointFast_ &&
    registry.SAVE.fastBuildOrUpdate === wmFastBuildOrUpdateLearningRecord_ &&
    registry.SAVE.fastStudyTimeFields === wmFastStudyTimeFields_ &&
    registry.SAVE.fastMergeStudyTimes === wmFastMergeStudyTimesIntoRecord_ &&
    registry.SAVE.fastApplyFinalTestOnly === wmFastApplyFinalTestOnly_ &&
    registry.SAVE.officialTimeFields === wmBuildOfficialLearningTimeFieldsV2_ &&
    registry.SAVE.buildRecord === WM_CODEGS_ALL_LOCK_V1.DB.buildRecord &&
    registry.SAVE.officialRound === WM_CODEGS_ALL_LOCK_V1.DB.officialRound &&
    registry.SAVE.findCurrentRecord === WM_CODEGS_ALL_LOCK_V1.DB.findCurrentRecord &&
    registry.SAVE.mergeCurrentRecord === WM_CODEGS_ALL_LOCK_V1.DB.mergeCurrentRecord &&
    registry.CURRENT_PROGRESS.upsert === WM_CODEGS_ALL_LOCK_V1.STUDY.currentProgress &&
    registry.CURRENT_PROGRESS.fields === WM_CODEGS_ALL_LOCK_V1.DB.currentProgressFields &&
    registry.CURRENT_PROGRESS.syncStudentSet === WM_CODEGS_ALL_LOCK_V1.DB.syncStudentCurrentSet &&
    registry.CURRENT_PROGRESS.clearCache === WM_CODEGS_ALL_LOCK_V1.DB.clearStudentCache &&
    registry.TRANSITION.evaluate === WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1.TRANSITION.evaluate &&
    registry.TRANSITION.levelCompletionState === WM_CODEGS_ALL_LOCK_V1.DB.levelCompletionState;

  return {
    version: registry.version,
    date: registry.date,
    modules: modules,
    registryFrozen: Object.isFrozen(registry),
    verified: allVerified
  };
}

/* WM ROUND RECORD RECOVERY DETAILED LOCK STATUS 20260727 */
function wmGetRoundRecordRecoveryLockStatus_() {
  var runtimeRegistry = WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1;
  var allRegistry = WM_CODEGS_ALL_LOCK_V1;
  var transitionRegistry = WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1;
  var rows = {
    getStudyProgressFast: {
      registered: runtimeRegistry.PROGRESS.fast === getStudyProgressFast,
      exactReference: allRegistry.STUDY.progressFast === getStudyProgressFast,
      registryFrozen: Object.isFrozen(runtimeRegistry.PROGRESS)
    },
    wmFindLatestIncompleteLearningRecordRowByStudentSet_: {
      registered: runtimeRegistry.SAVE.findCurrentRecord === wmFindLatestIncompleteLearningRecordRowByStudentSet_,
      exactReference: allRegistry.DB.findCurrentRecord === wmFindLatestIncompleteLearningRecordRowByStudentSet_,
      registryFrozen: Object.isFrozen(runtimeRegistry.SAVE)
    },
    wmMergeLatestIncompleteSetTimesForNewRecord_: {
      registered: runtimeRegistry.SAVE.mergeCurrentRecord === wmMergeLatestIncompleteSetTimesForNewRecord_,
      exactReference: allRegistry.DB.mergeCurrentRecord === wmMergeLatestIncompleteSetTimesForNewRecord_,
      registryFrozen: Object.isFrozen(runtimeRegistry.SAVE)
    },
    wmGetLockedLevelCompletionState_: {
      registered: runtimeRegistry.TRANSITION.levelCompletionState === wmGetLockedLevelCompletionState_,
      exactReference: transitionRegistry.TRANSITION.levelCompletionState === wmGetLockedLevelCompletionState_,
      registryFrozen: Object.isFrozen(runtimeRegistry.TRANSITION) && Object.isFrozen(transitionRegistry.TRANSITION)
    }
  };
  var verified = Object.keys(rows).every(function(name) {
    var row = rows[name];
    return row.registered === true && row.exactReference === true && row.registryFrozen === true;
  });
  return {
    title: 'WORD MATE ROUND RECORD RECOVERY LOCK STATUS',
    lockDate: '2026-07-27',
    functions: rows,
    verified: verified
  };
}


/* =========================================================
   🔒 WM TEST BOOTH FLOATING BANNER SERVER LOCK V1 / 실제 실행 연결
   TEST 전용 상단 플로팅 고정배너의 6개 화면 기준과
   F12 상태 주입 함수를 동결 Registry로 연결합니다.
   ========================================================= */
var WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_V1 = Object.freeze({
  version: 'WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_V1.0',
  date: '2026-07-28',
  environment: 'TEST',
  pages: Object.freeze(['LOGIN','MAPFRAME','LMS','DEV','MAP','STUDY']),
  buildConsole: wmBuildTestBoothFloatingBannerServerConsoleScript_,
  injectConsole: wmInjectTestBoothFloatingBannerServerConsole_
});

function wmGetTestBoothFloatingBannerServerLockStatus_() {
  var registry = WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_V1;
  var pages = registry.pages || [];
  var result = {
    version: registry.version,
    date: registry.date,
    environmentTest: WM_ENV === registry.environment,
    sixPagesRegistered: pages.length === 6 &&
      ['LOGIN','MAPFRAME','LMS','DEV','MAP','STUDY'].every(function(name){
        return pages.indexOf(name) !== -1;
      }),
    pagesFrozen: Object.isFrozen(pages),
    buildConsole: registry.buildConsole === wmBuildTestBoothFloatingBannerServerConsoleScript_,
    injectConsole: registry.injectConsole === wmInjectTestBoothFloatingBannerServerConsole_,
    registryFrozen: Object.isFrozen(registry)
  };
  result.verified = result.environmentTest &&
    result.sixPagesRegistered &&
    result.pagesFrozen &&
    result.buildConsole &&
    result.injectConsole &&
    result.registryFrozen;
  return result;
}

function wmBuildTestBoothFloatingBannerServerConsoleScript_(pageName) {
  var status = wmGetTestBoothFloatingBannerServerLockStatus_();
  status.page = String(pageName || '').trim().toUpperCase();
  return [
    '<script>',
    'window.WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_STATUS = Object.freeze(' + JSON.stringify(status) + ');',
    '(function(s){',
    'console.group("🔒 WORD MATE TEST BOOTH FLOATING BANNER SERVER LOCK STATUS — " + String(s.page || ""));',
    'console.log("LOCK VERSION : " + String(s.version || ""));',
    'console.log("LOCK DATE    : " + String(s.date || ""));',
    'console.table({',
    '  "TEST 환경":s.environmentTest,',
    '  "6개 화면 등록":s.sixPagesRegistered,',
    '  "화면목록 동결":s.pagesFrozen,',
    '  "F12 생성함수 실제 연결":s.buildConsole,',
    '  "F12 주입함수 실제 연결":s.injectConsole,',
    '  "Registry 동결":s.registryFrozen',
    '});',
    'console.log(s.verified ? "🔒 TEST BOOTH BANNER SERVER LOCK VERIFIED" : "❌ TEST BOOTH BANNER SERVER LOCK CHECK REQUIRED");',
    'console.groupEnd();',
    '})(window.WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_STATUS);',
    '</' + 'script>'
  ].join('\n');
}

function wmInjectTestBoothFloatingBannerServerConsole_(html, pageName) {
  html = String(html || '');
  var script = WM_TEST_BOOTH_FLOATING_BANNER_SERVER_LOCK_V1.buildConsole(pageName);
  return html.indexOf('</head>') !== -1
    ? html.replace('</head>', script + '\n</head>')
    : script + '\n' + html;
}

/* =========================================================
   🔒 WM CODE.GS ALL FUNCTION LOCK V1 / 실제 실행 연결
   운영 중인 로그인·세션·Map·Study·DB·LMS·설정센터 핵심 함수를
   변경 불가능한 Registry로 묶고 doGet 운영 경로에서 직접 실행합니다.
   ========================================================= */
const WM_CODEGS_ALL_LOCK_V1 = Object.freeze({
  version: 'WM_CODEGS_ALL_LOCK_V1.0',
  date: '2026-07-13',
  AUTH: Object.freeze({
    lmsLogin: wmLmsLogin,
    studentLogin: studentLogin,
    sessionCheck: checkStudentSession,
    studentLogout: studentLogout,
    sessionValid: wmIsStudentSessionValid_
  }),
  MAP: Object.freeze({
    studentProfile: getStudentBasicInfoForMapApi_,
    cache: getLearningMapCacheData,
    records: getLearningRecordsForMapData,
    fullMap: getLearningMap
  }),
  STUDY: Object.freeze({
    setData: getSetData,
    progressFast: getStudyProgressFast,
    progressLegacy: getLearningProgress,
    finalSave: wmSaveLearningRecordOfficialTotalV2_,
    currentProgress: wmUpsertCurrentProgress_
  }),
  DB: Object.freeze({
    buildRecord: wmBuildLearningRecordOfficialTotalV2_,
    officialRound: wmApplyOfficialCompleteRoundToRecord_,
    findCurrentRecord: wmFindLatestIncompleteLearningRecordRowByStudentSet_,
    mergeCurrentRecord: wmMergeLatestIncompleteSetTimesForNewRecord_,
    currentProgressFields: wmBuildCurrentProgressCacheFields_,
    levelCompletionState: wmGetLockedLevelCompletionState_,
    syncStudentCurrentSet: wmSyncStudentCurrentSetAfterOfficialCompletion_,
    clearStudentCache: wmClearRuntimeCachesForStudent_
  }),
  LMS: Object.freeze({
    students: getStudentsForLmsApi_,
    records: getLearningRecordsForLmsApi_,
    currentProgress: getCurrentProgressForLmsApi_,
    dashboard: getDashboardForLmsApi_
  }),
  STUDENT_ASSIGNMENT: Object.freeze({
    initialLevelSave: WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1.FIRST_ASSIGNMENT.save,
    currentSetRead: WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1.CURRENT_SET.read,
    nextLevelConditionSave: WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1.SAVE.studentAssignment,
    nextLevelPlusOne: WM_CODEGS_STUDENT_ASSIGNMENT_LOCK_V1.NEXT_LEVEL_PLUS_ONE.calculate
  }),
  SETTING: Object.freeze({
    getSettingCenter: wmGetSettingCenterForLms_,
    checkTeacherId: wmCheckSettingTeacherIdForLms_,
    saveTeacher: wmSaveSettingTeacherForLms_,
    saveTeacherStatus: wmSaveSettingTeacherStatusForLms_,
    saveClass: wmSaveSettingClassForLms_
  }),
  LMS_CORE: Object.freeze({
    settingGet: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.getSettingCenter,
    settingCheckTeacherId: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.checkTeacherId,
    settingSaveTeacher: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveTeacher,
    settingSaveTeacherStatus: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveTeacherStatus,
    settingSaveClass: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveClass,
    studentLoad: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.load,
    studentSave: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.save,
    studentCurrentSet: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.currentSet,
    studentActorFilter: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.actorFilter,
    recordLoad: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.load,
    recordSave: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.save,
    recordReadRows: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.readRows,
    recordActorFilter: WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.actorFilter
  }),
  TEST_BANNER: Object.freeze({
    status: wmGetTestBoothFloatingBannerServerLockStatus_,
    buildConsole: wmBuildTestBoothFloatingBannerServerConsoleScript_,
    injectConsole: wmInjectTestBoothFloatingBannerServerConsole_
  }),
  SYSTEM: Object.freeze({
    startLearning: startLearning,
    health: healthCheck,
    output: outputResult
  })
});

function wmGetCodeGsAllLockStatus_() {
  var registry = WM_CODEGS_ALL_LOCK_V1;
  var moduleNames = ['AUTH','MAP','STUDY','DB','LMS','STUDENT_ASSIGNMENT','SETTING','LMS_CORE','TEST_BANNER','SYSTEM'];
  var modules = {};
  var allVerified = Object.isFrozen(registry);

  moduleNames.forEach(function(moduleName) {
    var group = registry[moduleName] || {};
    var functions = {};
    var groupVerified = Object.isFrozen(group);
    Object.keys(group).forEach(function(functionName) {
      functions[functionName] = typeof group[functionName] === 'function';
      groupVerified = groupVerified && functions[functionName];
    });
    modules[moduleName] = { verified: groupVerified, functions: functions };
    allVerified = allVerified && groupVerified;
  });

  allVerified = allVerified &&
    registry.AUTH.studentLogin === WM_SPEED_LOCK_SERVER_V1.login &&
    registry.MAP.cache === WM_SPEED_LOCK_SERVER_V1.mapCache &&
    registry.MAP.records === WM_SPEED_LOCK_SERVER_V1.mapRecords &&
    registry.STUDY.setData === WM_STUDY_SPEED_SERVER_LOCK_V1.setDataFast &&
    registry.STUDY.progressFast === WM_STUDY_SPEED_SERVER_LOCK_V1.progressFast &&
    registry.STUDY.finalSave === WM_STUDY_SPEED_SERVER_LOCK_V1.finalSaveFast &&
    registry.STUDY.currentProgress === WM_STUDY_SPEED_SERVER_LOCK_V1.currentProgressFast &&
    registry.TEST_BANNER.status === wmGetTestBoothFloatingBannerServerLockStatus_ &&
    registry.TEST_BANNER.buildConsole === wmBuildTestBoothFloatingBannerServerConsoleScript_ &&
    registry.TEST_BANNER.injectConsole === wmInjectTestBoothFloatingBannerServerConsole_ &&
    wmGetTestBoothFloatingBannerServerLockStatus_().verified === true &&
    wmGetCodeGsLmsCoreActualLockStatus_().allVerified === true;

  return {
    version: registry.version,
    date: registry.date,
    modules: modules,
    registryFrozen: Object.isFrozen(registry),
    allVerified: allVerified
  };
}

function wmBuildCodeGsAllLockConsoleScript_() {
  var status = wmGetCodeGsAllLockStatus_();
  return [
    '<script>',
    'window.WM_CODEGS_ALL_LOCK_STATUS = Object.freeze(' + JSON.stringify(status) + ');',
    '(function(s){',
    'console.group("🔒 WORD MATE CODE.GS ALL LOCK STATUS");',
    'console.log("LOCK VERSION : " + String(s.version || ""));',
    'console.log("LOCK DATE    : " + String(s.date || ""));',
    'Object.keys(s.modules || {}).forEach(function(name){ console.log(name + " : " + (s.modules[name].verified ? "✅ ACTUAL LOCK" : "❌ CHECK REQUIRED"), s.modules[name].functions); });',
    'console.log(s.registryFrozen ? "✅ REGISTRY FROZEN" : "❌ REGISTRY CHECK REQUIRED");',
    'console.log(s.allVerified ? "🔒 CODE.GS ALL LOCK VERIFIED" : "❌ CODE.GS LOCK CHECK REQUIRED");',
    'console.groupEnd();',
    '})(window.WM_CODEGS_ALL_LOCK_STATUS);',
    '</' + 'script>'
  ].join('\n');
}

function wmInjectCodeGsAllLockConsole_(html) {
  html = String(html || '');
  var script = wmBuildCodeGsAllLockConsoleScript_();
  return html.indexOf('</head>') !== -1 ? html.replace('</head>', script + '\n</head>') : script + '\n' + html;
}

function wmBuildCodeGsLmsCoreActualLockConsoleScript_() {
  var status = wmGetCodeGsLmsCoreActualLockStatus_();
  return [
    '<script>',
    'window.WM_CODEGS_LMS_CORE_ACTUAL_LOCK_STATUS = Object.freeze(' + JSON.stringify(status) + ');',
    '(function(s){',
    'var rows = [];',
    'var total = 0;',
    'var locked = 0;',
    'Object.keys(s.modules || {}).forEach(function(moduleName){',
    '  var functions = (s.modules[moduleName] && s.modules[moduleName].functions) || {};',
    '  Object.keys(functions).forEach(function(functionName){',
    '    var ok = functions[functionName] === true;',
    '    total += 1;',
    '    if(ok) locked += 1;',
    '    rows.push({MODULE:moduleName, FUNCTION:functionName, LOCKED:ok});',
    '  });',
    '});',
    'console.group("🔒 WORD MATE CODE.GS LMS CORE ACTUAL LOCK STATUS");',
    'Object.keys(s.modules || {}).forEach(function(moduleName){',
    '  var functions = (s.modules[moduleName] && s.modules[moduleName].functions) || {};',
    '  var names = Object.keys(functions);',
    '  var count = names.filter(function(name){ return functions[name] === true; }).length;',
    '  console.log(moduleName + " : " + count + "/" + names.length);',
    '});',
    'console.table(rows);',
    'console.log("LMS_CORE : " + (s.allVerified ? "✅ ACTUAL LOCK" : "❌ CHECK REQUIRED"));',
    'console.log(locked + "/" + total + (locked === total && total === 13 ? " TRUE" : " CHECK REQUIRED"));',
    'console.log(s.registryFrozen ? "✅ REGISTRY FROZEN" : "❌ REGISTRY CHECK REQUIRED");',
    'console.groupEnd();',
    '})(window.WM_CODEGS_LMS_CORE_ACTUAL_LOCK_STATUS);',
    '</' + 'script>'
  ].join('\n');
}

function wmInjectCodeGsLmsCoreActualLockConsole_(html) {
  html = String(html || '');
  var script = wmBuildCodeGsLmsCoreActualLockConsoleScript_();
  return html.indexOf('</head>') !== -1 ? html.replace('</head>', script + '\n</head>') : script + '\n' + html;
}


/* WM_AUDIO_QA_STT_SERVER_V1 */
function wmAudioQaTranscribeForDev_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var audioBase64 = String(p.audioBase64 || '').trim();
  var expectedWord = String(p.expectedWord || '').trim();
  var languageCode = String(p.languageCode || 'en-US').trim() || 'en-US';

  if (!audioBase64) {
    return {success:false, message:'AUDIO_BASE64_REQUIRED'};
  }
  if (audioBase64.length > 9000000) {
    return {success:false, message:'AUDIO_PAYLOAD_TOO_LARGE'};
  }

  var payload = {
    config: {
      encoding: 'LINEAR16',
      languageCode: languageCode,
      enableAutomaticPunctuation: false,
      model: 'latest_short',
      useEnhanced: true,
      speechContexts: expectedWord ? [{phrases:[expectedWord], boost:20}] : []
    },
    audio: {
      content: audioBase64
    }
  };

  try {
    var response = UrlFetchApp.fetch(
      'https://speech.googleapis.com/v1/speech:recognize',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      }
    );

    var status = response.getResponseCode();
    var body = response.getContentText() || '{}';
    var data = JSON.parse(body);

    if (status < 200 || status >= 300) {
      var apiMessage = data &&
        data.error &&
        data.error.message
        ? data.error.message
        : 'SPEECH_API_HTTP_' + status;

      return {
        success:false,
        message:apiMessage,
        apiStatus:status
      };
    }

    var alternatives = [];
    (data.results || []).forEach(function(result) {
      (result.alternatives || []).forEach(function(alt) {
        alternatives.push(alt);
      });
    });

    alternatives.sort(function(a, b) {
      return Number(b.confidence || 0) - Number(a.confidence || 0);
    });

    var best = alternatives[0] || {};

    return {
      success:true,
      transcript:String(best.transcript || '').trim(),
      confidence:Number(best.confidence || 0),
      expectedWord:expectedWord
    };
  } catch (err) {
    return {
      success:false,
      message:err && err.message ? err.message : String(err)
    };
  }
}

var WM_AUDIO_QA_STT_SERVER_LOCK_V1 = Object.freeze({
  version:'WM_AUDIO_QA_STT_SERVER_V1.2',
  date:'2026-08-07',
  TRANSCRIBE:Object.freeze({
    run:wmAudioQaTranscribeForDev_
  }),
  STATUS:Object.freeze({
    read:wmGetAudioQaSttServerLockStatus_
  })
});

function wmGetAudioQaSttServerLockStatus_() {
  var registry = WM_AUDIO_QA_STT_SERVER_LOCK_V1;
  var result = {
    version:registry.version,
    date:registry.date,
    registryFrozen:Object.isFrozen(registry),
    transcribeGroupFrozen:Object.isFrozen(registry.TRANSCRIBE),
    statusGroupFrozen:Object.isFrozen(registry.STATUS),
    transcribeExists:typeof registry.TRANSCRIBE.run === 'function',
    statusExists:typeof registry.STATUS.read === 'function',
    transcribeExact:registry.TRANSCRIBE.run === wmAudioQaTranscribeForDev_,
    statusExact:registry.STATUS.read === wmGetAudioQaSttServerLockStatus_
  };
  result.verified = result.registryFrozen &&
    result.transcribeGroupFrozen &&
    result.statusGroupFrozen &&
    result.transcribeExists &&
    result.statusExists &&
    result.transcribeExact &&
    result.statusExact;
  return result;
}


function wmIsAudioQaProtectedRequest_(action, mode) {
  return action === 'audioQaTranscribe' || mode === 'audioQaTranscribe' ||
    action === 'audioQaSttLockStatus' || mode === 'audioQaSttLockStatus';
}

function doGet(e) {
  try {
    var mode = '';
    if (e && e.parameter && e.parameter.mode) {
      mode = String(e.parameter.mode).trim();
    }

    var action = '';
    if (e && e.parameter && e.parameter.action) {
      action = String(e.parameter.action).trim();
    }

    if (wmIsAudioQaProtectedRequest_(action, mode) && !wmGetAudioQaSttServerLockStatus_().verified) {
      return outputResult(e, {success:false, lockBlocked:true, message:'AUDIO QA STT SERVER LOCK VERIFY FAILED'});
    }

    if (action === 'audioQaTranscribe' || mode === 'audioQaTranscribe') {
      return outputResult(e, WM_AUDIO_QA_STT_SERVER_LOCK_V1.TRANSCRIBE.run(e));
    }

    if (action === 'audioQaSttLockStatus' || mode === 'audioQaSttLockStatus') {
      return outputResult(e, wmGetAudioQaSttServerLockStatus_());
    }

    if (wmIsCodeGsLmsCoreProtectedRequest_(action, mode) && !wmGetCodeGsLmsCoreActualLockStatus_().allVerified) {
      return outputResult(e, {success:false, lockBlocked:true, message:'CODE.GS LMS CORE FUNCTION LOCK VERIFY FAILED'});
    }

    if (action === 'lmsLogin' || action === 'wmLmsLogin' || mode === 'lmsLogin' || mode === 'wmLmsLogin') {
      var loginParams = (e && e.parameter) ? e.parameter : {};
      return outputResult(e, WM_CODEGS_ALL_LOCK_V1.AUTH.lmsLogin(
        loginParams.teacherId || loginParams.teacherID || loginParams.id || '',
        loginParams.password || loginParams.pw || ''
      ));
    }

    if (action === 'saveStudyCheckpointFast') {
      return WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.SAVE.checkpointFast(e);
    }

    if (action === 'saveLearningRecordV2') {
      return WM_CODEGS_ALL_LOCK_V1.STUDY.finalSave(e);
    }

    if (action === 'saveLearningRecord') {
      return WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.save(e);
    }


    if (action === 'getStudentBasicInfoForMap' || mode === 'getStudentBasicInfoForMap') {
      return WM_CODEGS_ALL_LOCK_V1.MAP.studentProfile(e);
    }

    if (action === 'getStudentLearningMode' || mode === 'getStudentLearningMode') {
      return WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.LEARNING_MODE_API.read(e);
    }

    if (action === 'saveStudentLearningMode' || mode === 'saveStudentLearningMode') {
      return WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.save(e);
    }

    if (action === 'students' || mode === 'students') {
      return WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDENT_DB.load(e);
    }

    if (action === 'records' || mode === 'records' || action === 'learningRecords' || mode === 'learningRecords') {
      return WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.STUDY_RECORD_DB.load(e);
    }

    /* WM_LMS_CUMULATIVE_WRONG_V1
     * 구글시트 컬럼 추가 없이 학생별 최신레벨 누적오답을 클릭 시점에만 계산합니다. */
    if (action === 'cumulativeWrong' || mode === 'cumulativeWrong') {
      return wmGetCumulativeWrongForLmsApi_(e);
    }

    /* WM_GRADE_LEVEL_DB_LMS_API_ROUTE_V1
     * 6.성적등급 탭 진입 시 6.성적등급_DB를 계산·변형 없이 그대로 반환합니다.
     */
    if (action === 'gradeLevels' || mode === 'gradeLevels' || action === 'gradeLevel' || mode === 'gradeLevel') {
      return wmGetGradeLevelsForLmsApi_(e);
    }

    if (action === 'reports' || mode === 'reports') {
      return wmGetReportsForLmsApi_(e);
    }

    if (action === 'monthlyReportData' || mode === 'monthlyReportData') {
      return outputResult(e, wmGetMonthlyReportDataForLms_(e));
    }

    if (action === 'recentReportMonths' || mode === 'recentReportMonths') {
      return outputResult(e, wmGetRecentReportMonthsForLms_());
    }

    if (action === 'publicReport' || mode === 'publicReport') {
      if (!wmGetCodeGsLmsPhase2ReportGradeLockStatus_().allVerified) {
        return outputResult(e, {success:false, lockBlocked:true, message:'성적표를 확인할 수 없습니다.', data:null});
      }
      return outputResult(e, wmGetPublicReportByToken_(e));
    }

    if (action === 'initializeReportsFromLearningRecords' || mode === 'initializeReportsFromLearningRecords') {
      return outputResult(e, wmInitializeReportsFromLearningRecordsForLms_(e));
    }

    if (action === 'currentProgress' || mode === 'currentProgress') {
      return WM_CODEGS_ALL_LOCK_V1.LMS.currentProgress(e);
    }

    /* WM_LMS_DASHBOARD_API_ROUTE_20260706_V1
     * LMS.html 대시보드가 호출하는 ?mode=dashboard 응답을 제공합니다.
     * 학생관리_DB / 학습기록_DB / 현재진행_DB를 그대로 읽어 대시보드 집계 소스로 반환합니다.
     */
    if (action === 'dashboard' || mode === 'dashboard') {
      return WM_CODEGS_ALL_LOCK_V1.LMS.dashboard(e);
    }

    if (action === 'sendCenterData' || mode === 'sendCenterData') {
      return outputResult(e, wmGetSendCenterDataForLms_(e));
    }

    if (action === 'saveSendSetting' || mode === 'saveSendSetting') {
      return outputResult(e, wmSaveSendSettingForLms_(e));
    }

    if (action === 'getSettingCenter' || mode === 'getSettingCenter') {
      return outputResult(e, WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.getSettingCenter(e));
    }

    if (action === 'checkTeacherId' || mode === 'checkTeacherId') {
      return outputResult(e, WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.checkTeacherId(e));
    }

    if (action === 'saveSettingTeacher' || mode === 'saveSettingTeacher') {
      return outputResult(e, WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveTeacher(e));
    }

    if (action === 'saveSettingTeacherStatus' || mode === 'saveSettingTeacherStatus') {
      return outputResult(e, WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveTeacherStatus(e));
    }

    if (action === 'saveSettingTeacherDataAccess' || mode === 'saveSettingTeacherDataAccess') {
      return outputResult(e, wmSaveSettingTeacherDataAccessForLms_(e));
    }

    if (action === 'saveSettingClass' || mode === 'saveSettingClass') {
      return outputResult(e, WM_CODEGS_LMS_CORE_ACTUAL_LOCK_V1.SETTING_CENTER.saveClass(e));
    }

    if (action === 'saveSettingClassStatus' || mode === 'saveSettingClassStatus') {
      return outputResult(e, wmSaveSettingClassStatusForLms_(e));
    }


    if (action === 'generateStudentId' || mode === 'generateStudentId') {
      return outputResult(e, wmGenerateStudentIdForLms_());
    }


    if (action === 'checkStudentId' || mode === 'checkStudentId') {
      return outputResult(e, wmCheckStudentIdForLms_(e));
    }

    if (action === 'saveStudentToDB' || mode === 'saveStudentToDB') {
      return outputResult(e, wmSaveStudentToDbForLms_(e));
    }

    if (action === 'deleteStudent' || mode === 'deleteStudent') {
      return outputResult(e, wmDeleteStudentForLms_(e));
    }

    if (action === 'deletedStudents' || mode === 'deletedStudents') {
      return outputResult(e, wmGetDeletedStudentsForLms_(e));
    }

    if (action === 'checkStudentSession' || mode === 'checkStudentSession') {
      return WM_CODEGS_ALL_LOCK_V1.AUTH.sessionCheck(e);
    }

    if (action === 'studentLogout' || mode === 'studentLogout') {
      return WM_CODEGS_ALL_LOCK_V1.AUTH.studentLogout(e);
    }

    if (action === 'speedLockStatus' || mode === 'speedLockStatus') {
      return outputResult(e, wmGetSpeedLockServerStatus_());
    }

    if (action === 'studySpeedLockStatus' || mode === 'studySpeedLockStatus') {
      return outputResult(e, wmGetStudySpeedServerLockStatus_());
    }

    if (action === 'studyRuntimeExecutionLockStatus' || mode === 'studyRuntimeExecutionLockStatus') {
      return outputResult(e, wmGetStudyRuntimeExecutionServerLockStatus_());
    }

    if (action === 'roundRecordRecoveryLockStatus' || mode === 'roundRecordRecoveryLockStatus') {
      return outputResult(e, wmGetRoundRecordRecoveryLockStatus_());
    }

    if (action === 'settingCenterLockStatus' || mode === 'settingCenterLockStatus') {
      return outputResult(e, wmGetSettingCenterServerLockStatus_());
    }

    if (action === 'codeGsAllLockStatus' || mode === 'codeGsAllLockStatus') {
      return outputResult(e, wmGetCodeGsAllLockStatus_());
    }

    if (action === 'testBoothFloatingBannerLockStatus' || mode === 'testBoothFloatingBannerLockStatus') {
      return outputResult(e, WM_CODEGS_ALL_LOCK_V1.TEST_BANNER.status());
    }

    if (action === 'codeGsStudentAssignmentLockStatus' || mode === 'codeGsStudentAssignmentLockStatus') {
      return outputResult(e, wmGetCodeGsStudentAssignmentLockStatus_());
    }

    if (action === 'levelCompleteConditionLockStatus' || mode === 'levelCompleteConditionLockStatus') {
      return outputResult(e, wmGetCodeGsLevelCompleteConditionLockStatus_());
    }

    if (action === 'codeGsLmsCoreActualLockStatus' || mode === 'codeGsLmsCoreActualLockStatus') {
      return outputResult(e, wmGetCodeGsLmsCoreActualLockStatus_());
    }

    if (action === 'levelTransitionNoticeLockStatus' || mode === 'levelTransitionNoticeLockStatus') {
      return outputResult(e, wmGetLevelTransitionNoticeServerLockStatus_());
    }

    if (action === 'getSetData') {
      return WM_CODEGS_ALL_LOCK_V1.STUDY.setData(e);
    }

    if (action === 'getLearningProgress' || mode === 'getLearningProgress') {
      return WM_CODEGS_ALL_LOCK_V1.STUDY.progressLegacy(e);
    }

    if (action === 'getStudyProgressFast' || mode === 'getStudyProgressFast') {
      return WM_CODEGS_ALL_LOCK_V1.STUDY.progressFast(e);
    }

    /* WM_GET_LEARNING_MAP_ROUTE_FIRST_V2
     * 학습맵 하단 숙제검사는 JSON/API 응답을 받아야 하므로
     * 화면 라우터(integrated/map/study)보다 먼저 처리합니다.
     */
    if (action === 'getLearningMapCacheData' || mode === 'getLearningMapCacheData') {
      var cacheStudentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid)) || '').trim();
      var cacheSessionToken = String((e && e.parameter && (e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN)) || '').trim();
      if (cacheSessionToken && !wmIsStudentSessionValid_(cacheStudentId, cacheSessionToken)) {
        return outputResult(e, wmBuildSessionExpiredResponse_());
      }
      var cachePayload = WM_CODEGS_ALL_LOCK_V1.MAP.cache(cacheStudentId);
      if (cachePayload && cacheSessionToken) {
        cachePayload.sessionValid = true;
        cachePayload.valid = true;
      }
      return outputResult(e, cachePayload);
    }

    if (action === 'getLearningRecordsForMapData' || mode === 'getLearningRecordsForMapData') {
      var recordsStudentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid)) || '').trim();
      var recordsPage = Number((e && e.parameter && e.parameter.page) || 0);
      var recordsPageSize = Number((e && e.parameter && e.parameter.pageSize) || 0);
      var recordsFromDate = String((e && e.parameter && e.parameter.fromDate) || '').trim();
      var recordsToDate = String((e && e.parameter && e.parameter.toDate) || '').trim();
      var recordsSearchFilter = String((e && e.parameter && e.parameter.searchFilter) || '').trim();
      return outputResult(e, WM_CODEGS_ALL_LOCK_V1.MAP.records(
        recordsStudentId,
        recordsPage,
        recordsPageSize,
        recordsFromDate,
        recordsToDate,
        recordsSearchFilter
      ));
    }

    if (mode === 'getLearningMap') {
      return WM_CODEGS_ALL_LOCK_V1.MAP.fullMap(e);
    }

    /* WM_INTEGRATED_DIRECT_SEPARATED_ROUTE_V1
     * 통합 URL로 들어와도 현재 오픈 우선 흐름은 분리 파일(Index/Map/Study/Dev)을 직접 실행합니다.
     * Integrated.html router 화면에서 멈추는 문제를 피하고 같은 /exec 기준으로 분리 운영합니다.
     */
    if (mode === 'integrated' || mode === 'integratedIndex' || mode === 'integratedMap' || mode === 'integratedDev' || mode === 'integratedStudy') {
      return renderIntegratedDirectSeparatedApp_(e, mode);
    }
if (mode === 'lms') {
  var lmsHtml = HtmlService.createHtmlOutputFromFile('Index').getContent();
  lmsHtml = wmInjectCodeGsAllLockConsole_(lmsHtml);
  lmsHtml = wmInjectCodeGsLmsCoreActualLockConsole_(lmsHtml);
  lmsHtml = WM_CODEGS_ALL_LOCK_V1.TEST_BANNER.injectConsole(lmsHtml, 'LMS');
  return HtmlService
    .createHtmlOutput(lmsHtml)
    .setTitle('Word Mate LMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
    if (mode === 'map') {
      var mapStudentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid || e.parameter.wmStudentId)) || '').trim();
      var mapStudentName = String((e && e.parameter && (e.parameter.studentName || e.parameter.name || e.parameter.wmStudentName || e.parameter.wmName)) || '').trim();
      var mapLearningAssign = String((e && e.parameter && (e.parameter.learningAssign || e.parameter.level || e.parameter.learningLevel || e.parameter.wmLearningAssign)) || '').trim();
      var mapSetId = String((e && e.parameter && (e.parameter.setId || e.parameter.set_id || e.parameter.Set_ID || e.parameter.currentSetId || e.parameter.currentSet || e.parameter.wmCurrentSet)) || '').trim();
      var mapSessionToken = String((e && e.parameter && (e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN)) || '').trim();

      var mapPayload = parseMapStudentPayload_(e && e.parameter ? (e.parameter.wmStudentPayload || e.parameter.studentPayload || '') : '');
      if (mapPayload) {
        if (!mapStudentId) {
          mapStudentId = String(mapPayload['학생ID'] || mapPayload.studentId || '').trim();
        }
        if (!mapStudentName) {
          mapStudentName = String(mapPayload['학생이름'] || mapPayload.studentName || mapPayload.name || '').trim();
        }
        if (!mapLearningAssign) {
          mapLearningAssign = String(mapPayload['학습배정'] || mapPayload.learningAssign || mapPayload.level || '').trim();
        }
        if (!mapSetId) {
          mapSetId = String(mapPayload['현재세트'] || mapPayload.currentSet || '').trim();
        }
        if (!mapSessionToken) {
          mapSessionToken = String(mapPayload['현재세션'] || mapPayload.sessionToken || mapPayload.wmSessionToken || '').trim();
        }
      }

      /* WM_MAP_STUDENT_DB_INJECT_FIX_V8
       * 로그인창에서 studentName이 전달되지 않거나 HtmlService iframe에서 URL 파라미터가 끊겨도,
       * studentId만 있으면 1.학생관리_DB에서 학생이름/학습배정을 다시 조회해 Map.html에 직접 주입합니다.
       */
      var mapStudentInfo = getStudentBasicInfoForMap_(mapStudentId);
      var mapCurrentProgress = wmGetCurrentProgressCacheForStudent_(mapStudentId) || null;
      var mapCurrentLevel = String(mapCurrentProgress && mapCurrentProgress['현재레벨'] || '').trim();
      var mapProgressCurrentSet = mapCurrentProgress
        ? wmNormalizeCurrentProgressSetIdForMap_(mapCurrentProgress, '')
        : '';

      if (!mapStudentName && mapStudentInfo.studentName) {
        mapStudentName = mapStudentInfo.studentName;
      }

      if (!mapLearningAssign && mapStudentInfo.learningAssign) {
        mapLearningAssign = mapStudentInfo.learningAssign;
      }

      if (mapProgressCurrentSet) {
        mapSetId = mapProgressCurrentSet;
      } else {
        mapSetId = '';
      }

      /* WM_MAP_API_EXEC_URL_INJECT_V1
       * Map.html 안에서 window.location.href를 API 주소로 오인하지 않도록
       * 현재 배포된 Apps Script /exec 주소를 서버에서 직접 주입합니다.
       */
      var mapExecUrl = '';
      try {
        mapExecUrl = ScriptApp.getService().getUrl();
      } catch (mapExecUrlErr) {
        mapExecUrl = '';
      }

      var mapStudentPayload = {
        학생ID: mapStudentId,
        학생이름: mapStudentName,
        학교: mapStudentInfo.school || '',
        학년: mapStudentInfo.grade || '',
        Class: mapStudentInfo.className || '',
        교사명: mapStudentInfo.teacherName || '',
        학습배정: mapLearningAssign,
        현재레벨: mapCurrentLevel,
        현재세트: mapSetId,
        현재세션: mapSessionToken,
        sessionToken: mapSessionToken,
        학습모드: getStudentLearningMode_(mapStudentId, '')
      };

      var mapHtml = HtmlService
        .createHtmlOutputFromFile('Map.html')
        .getContent();
      var mapLoginSpeedLockStatus = wmGetSpeedLockServerStatus_();

      var mapBootScript = [
        '<script>',
        'window.WM_MAP_SERVER_STUDENT = ' + JSON.stringify(mapStudentPayload) + ';',
        'window.WM_LOGGED_IN_STUDENT = ' + JSON.stringify(mapStudentPayload) + ';',
        'window.WM_CURRENT_STUDENT = ' + JSON.stringify(mapStudentPayload) + ';',
        'window.WM_LOGGED_IN_STUDENT_ID = ' + JSON.stringify(mapStudentId) + ';',
        'window.WM_STUDENT_ID = ' + JSON.stringify(mapStudentId) + ';',
        'window.WM_STUDENT_NAME = ' + JSON.stringify(mapStudentName) + ';',
        'window.WM_LEARNING_ASSIGN = ' + JSON.stringify(mapLearningAssign) + ';',
        'window.WM_INITIAL_SET_ID = ' + JSON.stringify(mapSetId) + ';',
        'window.WM_SESSION_TOKEN = ' + JSON.stringify(mapSessionToken) + ';',
        'window.WM_APPS_SCRIPT_URL = ' + JSON.stringify(mapExecUrl) + ';',
        'window.WM_EXEC_URL = ' + JSON.stringify(mapExecUrl) + ';',
        'window.WM_LOGIN_SERVER_SPEED_LOCK = Object.freeze(' + JSON.stringify(mapLoginSpeedLockStatus) + ');',
        'console.group("🔒 WORD MATE LOGIN SERVER SPEED LOCK STATUS");',
        'console.log("LOCK VERSION : " + String(window.WM_LOGIN_SERVER_SPEED_LOCK.version || ""));',
        'console.log("LOGIN ROUTE  : " + (window.WM_LOGIN_SERVER_SPEED_LOCK.login === true ? "✅ ACTUAL LOCK" : "❌ CHECK REQUIRED"));',
        'console.log("OPTIMIZATION : " + (window.WM_LOGIN_SERVER_SPEED_LOCK.loginOptimization === true ? "✅ ACTUAL LOCK" : "❌ CHECK REQUIRED"));',
        'console.log(window.WM_LOGIN_SERVER_SPEED_LOCK.loginVerified === true ? "🔒 LOGIN SERVER SPEED LOCK VERIFIED" : "⚠️ LOGIN SERVER SPEED LOCK CHECK REQUIRED");',
        'console.groupEnd();',
        'try {',
        '  sessionStorage.setItem("wmLoggedInStudent", ' + JSON.stringify(JSON.stringify(mapStudentPayload)) + ');',
        '  sessionStorage.setItem("WM_STUDENT_ID", ' + JSON.stringify(mapStudentId) + ');',
        '  sessionStorage.setItem("WM_STUDENT_NAME", ' + JSON.stringify(mapStudentName) + ');',
        '  sessionStorage.setItem("WM_LEARNING_ASSIGN", ' + JSON.stringify(mapLearningAssign) + ');',
        '  sessionStorage.setItem("WM_INITIAL_SET_ID", ' + JSON.stringify(mapSetId) + ');',
        '  sessionStorage.setItem("WM_SESSION_TOKEN", ' + JSON.stringify(mapSessionToken) + ');',
        '  sessionStorage.setItem("WM_APPS_SCRIPT_URL", ' + JSON.stringify(mapExecUrl) + ');',
        '} catch (e) {}',
        '</' + 'script>'
      ].join('\n');

      mapHtml = mapHtml
        .replace(/__WM_STUDENT_ID__/g, mapStudentId)
        .replace(/__WM_STUDENT_NAME__/g, mapStudentName)
        .replace(/__WM_LEARNING_ASSIGN__/g, mapLearningAssign)
        .replace(/__WM_INITIAL_SET_ID__/g, mapSetId);

      /* WM_MAP_ROOT_DATA_FIX_V13
       * Map.html 초기화 JS가 다시 실행되어도 학생정보가 사라지지 않도록
       * 루트 DOM에 학생정보를 data-*로 직접 심어 둡니다.
       */
      mapHtml = mapHtml.replace(
        /<div id=["']wm-map-root["']>/,
        '<div id="wm-map-root"'
          + ' data-student-id="' + escapeHtmlForMap_(mapStudentId) + '"'
          + ' data-student-name="' + escapeHtmlForMap_(mapStudentName) + '"'
          + ' data-learning-assign="' + escapeHtmlForMap_(mapLearningAssign) + '"'
          + ' data-current-set="' + escapeHtmlForMap_(mapSetId) + '">'
      );

      /* WM_MAP_SERVER_DIRECT_TEXT_FIX_V10
       * JS 전달이 실패해도 첫 화면 상단 학생명은 서버에서 직접 박아 넣습니다.
       */
      if (mapStudentName || mapStudentId) {
        mapHtml = mapHtml.replace(
          /<strong id=["']studentName["']>.*?<\/strong>/,
          '<strong id="studentName">' + escapeHtmlForMap_(mapStudentName || mapStudentId) + '</strong>'
        );
      }

      if (mapCurrentLevel) {
        var directLevelText = extractLevelNumberForMap_(mapCurrentLevel) + ' Level';
        mapHtml = mapHtml.replace(
          /<strong id=["']studentLevel["']>.*?<\/strong>/,
          '<strong id="studentLevel">' + escapeHtmlForMap_(directLevelText) + '</strong>'
        );
      }

      if (mapHtml.indexOf('</head>') !== -1) {
        mapHtml = mapHtml.replace('</head>', mapBootScript + '\n</head>');
      } else {
        mapHtml = mapBootScript + '\n' + mapHtml;
      }
      mapHtml = wmInjectCodeGsAllLockConsole_(mapHtml);
      mapHtml = WM_CODEGS_ALL_LOCK_V1.TEST_BANNER.injectConsole(mapHtml, 'MAP');

      return HtmlService
        .createHtmlOutput(mapHtml)
        .setTitle('Word Mate 학습맵')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'dev') {
      var devHtml = HtmlService.createHtmlOutputFromFile('Dev.html').getContent();
      devHtml = wmInjectCodeGsAllLockConsole_(devHtml);
      devHtml = WM_CODEGS_ALL_LOCK_V1.TEST_BANNER.injectConsole(devHtml, 'DEV');
      return HtmlService
        .createHtmlOutput(devHtml)
        .setTitle('Word Mate 학습화면');
    }

    if (mode === 'study') {

      var initialSetId =
        String(
          e.parameter.set_id ||
          e.parameter.setId ||
          e.parameter.Set_ID ||
          e.parameter.currentSetId ||
          'WM5-1-1'
        )
        .trim()
        .toUpperCase();

      if (!/^WM\d+-\d+-\d+$/.test(initialSetId)) {
        initialSetId = 'WM5-1-1';
      }

      var studentId =
        String(
          e.parameter.studentId ||
          e.parameter.studentID ||
          e.parameter.sid ||
          ''
        )
        .trim()
        .toUpperCase();

      var studentName =
        String(
          e.parameter.studentName ||
          e.parameter.name ||
          ''
        )
        .trim();

      var studySessionToken = String(e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN || '').trim();
      var studyStudentInfo = WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.BOOT_CONTEXT.studentInfo(studentId);
      if (!studentName && studyStudentInfo.studentName) {
        studentName = studyStudentInfo.studentName;
      }
      var studyLearningMode = WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.BOOT_CONTEXT.learningMode(studentId, '');
      var studyExecUrl = '';
      try {
        studyExecUrl = ScriptApp.getService().getUrl();
      } catch (studyExecUrlErr) {
        studyExecUrl = '';
      }

      var studyHtml =
        HtmlService
          .createHtmlOutputFromFile('Study')
          .getContent();

      var studyBootScript = [
        '<script>',
        'window.WM_STUDY_MODE_FROM_SERVER = ' + JSON.stringify(studyLearningMode) + ';',
        'window.WM_LEARNING_ASSIGN_FROM_SERVER = ' + JSON.stringify(studyStudentInfo.learningAssign || '') + ';',
        'window.WM_APPS_SCRIPT_URL = ' + JSON.stringify(studyExecUrl) + ';',
        'window.WM_EXEC_URL = ' + JSON.stringify(studyExecUrl) + ';',
        '</' + 'script>'
      ].join('\n');

      studyHtml =
        studyHtml
          .replace(/__WM_INITIAL_SET_ID__/g, initialSetId)
          .replace(/__WM_STUDENT_ID__/g, studentId)
          .replace(/__WM_STUDENT_NAME__/g, studentName)
          .replace(/__WM_SESSION_TOKEN__/g, studySessionToken);

      if (studyHtml.indexOf('</head>') !== -1) {
        studyHtml = studyHtml.replace('</head>', studyBootScript + '\n</head>');
      } else {
        studyHtml = studyBootScript + '\n' + studyHtml;
      }
      studyHtml = wmInjectCodeGsAllLockConsole_(studyHtml);
      studyHtml = WM_CODEGS_ALL_LOCK_V1.TEST_BANNER.injectConsole(studyHtml, 'STUDY');

      return HtmlService
        .createHtmlOutput(studyHtml)
        .setTitle('Word Mate Study');
    }

    if (mode === 'studentLogin') {
      return WM_CODEGS_ALL_LOCK_V1.AUTH.studentLogin(e);
    }

    if (mode === 'startLearning') {
      return WM_CODEGS_ALL_LOCK_V1.SYSTEM.startLearning(e);
    }


    if (mode === 'wmDebugCheck') {
      return wmDebugCheck(e);
    }

    if (mode === 'versionCheck') {
  return outputResult(e, {
    success: true,
    message: '현재 /exec가 Integrated Router V7 대응 Code.gs를 보고 있습니다.'
  });
}

    if (mode === 'health') {
      return WM_CODEGS_ALL_LOCK_V1.SYSTEM.health(e);
    }

    return outputResult(e, {
      success: false,
      message: 'no mode'
    });

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: 'Apps Script 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}


/* WM_LMS_TEACHER_LOGIN_AUTH_V1
 * /lms 진입 전 7-1.교사관리_DB의 교사ID + 비밀번호를 검증합니다.
 * 기존 학생/학습 로직에는 영향이 없도록 독립 함수로만 추가합니다.
 */
function wmLmsLogin(teacherId, password) {
  var inputId = String(teacherId || '').trim();
  var inputIdKey = inputId.toLowerCase();
  var inputPw = String(password || '').trim();

  if (!inputId || !inputPw) {
    return { success:false, message:'교사ID와 비밀번호를 입력하세요.' };
  }

  if (!/^[A-Za-z0-9]{1,10}$/.test(inputId)) {
    return { success:false, message:'교사ID는 영문/숫자 10자리 이내입니다.' };
  }

  if (!/^\d{4}$/.test(inputPw)) {
    return { success:false, message:'비밀번호는 숫자 4자리입니다.' };
  }

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = wmGetTeacherAccountSheet_(ss);
    if (!sheet) {
      return { success:false, message:'7-1.교사관리_DB 시트를 찾을 수 없습니다.' };
    }

    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) {
      return { success:false, message:'교사 계정 데이터가 없습니다.' };
    }

    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID','Teacher_ID','TeacherID','teacherId']);
    var idxName = wmFindHeaderIndex_(headers, ['교사명','이름','담당교사','name']);
    var idxPassword = wmFindHeaderIndex_(headers, ['비밀번호','비번','패스워드','password','Password']);
    var idxRole = wmFindHeaderIndex_(headers, ['교사등급','권한','role','Role']);
    var idxLeader = wmFindHeaderIndex_(headers, ['담당리더','리더']);
    var idxTeacher = wmFindHeaderIndex_(headers, ['담당교사','교사']);
    var idxClass = wmFindHeaderIndex_(headers, ['담당반목록','담당반','반']);
    var idxAccessLevel = wmFindHeaderIndex_(headers, ['접근레벨']);
    var idxParentAccess = wmFindHeaderIndex_(headers, ['학부모정보권한','학부모관련권한']);
    var idxPaymentAccess = wmFindHeaderIndex_(headers, ['결제정보권한','결제관련권한']);
    var idxStatus = wmFindHeaderIndex_(headers, ['교사상태','상태','활성상태']);
    var idxRecentLogin = wmFindHeaderIndex_(headers, ['최근로그인','최종로그인']);
    var idxUpdatedAt = wmFindHeaderIndex_(headers, ['수정일','수정일시']);

    if (idxTeacherId < 0) {
      return { success:false, message:'교사ID 컬럼을 찾을 수 없습니다.' };
    }
    if (idxPassword < 0) {
      return { success:false, message:'비밀번호 컬럼을 찾을 수 없습니다. 7-1.교사관리_DB에 비밀번호 컬럼을 추가하세요.' };
    }

    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var rowTeacherId = String(row[idxTeacherId] || '').trim();
      if (rowTeacherId.toLowerCase() !== inputIdKey) continue;

      var rowPassword = String(row[idxPassword] || '').trim();
      if (rowPassword !== inputPw) {
        return { success:false, message:'비밀번호가 맞지 않습니다.' };
      }

      var teacherStatus = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : '사용';
      if (teacherStatus && !/^(사용|활성|정상|ON|Y|TRUE|ACTIVE|SUPER ADMIN|ADMIN)$/i.test(teacherStatus)) {
        return { success:false, message:'사용 중지된 교사 계정입니다.' };
      }

      /* 로그인 속도 개선: 최근로그인만 기록하고 일반 수정일은 로그인마다 갱신하지 않습니다. */
      if (idxRecentLogin >= 0) {
        try { sheet.getRange(r + 1, idxRecentLogin + 1).setValue(new Date()); } catch (ignoreRecentLoginWrite) {}
      }

      var rawRole = idxRole >= 0 ? String(row[idxRole] || '').trim() : 'TEACHER';
      var normalizedRole = wmNormalizeLmsRole_(rawRole);

      return {
        success:true,
        message:'LMS 로그인 성공',
        teacher:{
          교사ID: rowTeacherId,
          이름: idxName >= 0 ? String(row[idxName] || '').trim() : rowTeacherId,
          권한: normalizedRole,
          원본권한: rawRole,
          담당리더: idxLeader >= 0 ? String(row[idxLeader] || '').trim() : '',
          담당교사: idxTeacher >= 0 ? String(row[idxTeacher] || '').trim() : '',
          담당반: idxClass >= 0 ? String(row[idxClass] || '').trim() : '',
          접근레벨: idxAccessLevel >= 0 ? String(row[idxAccessLevel] || '').trim() : '',
          학부모정보권한: wmNormalizeStudentDataAccessFlag_(idxParentAccess >= 0 ? row[idxParentAccess] : 'OFF'),
          결제정보권한: wmNormalizeStudentDataAccessFlag_(idxPaymentAccess >= 0 ? row[idxPaymentAccess] : 'OFF')
        }
      };
    }

    return { success:false, message:'등록된 교사ID가 없습니다.' };
  } catch (err) {
    return { success:false, message:'LMS 로그인 검증 오류', error:String(err && err.message ? err.message : err) };
  }
}

/* WM_LMS_TEACHER_LOGIN_ALIAS_V1 */
function teacherLogin(teacherId, password) {
  return wmLmsLogin(teacherId, password);
}

function checkTeacherLogin(teacherId, password) {
  return wmLmsLogin(teacherId, password);
}

function wmTeacherLogin(teacherId, password) {
  return wmLmsLogin(teacherId, password);
}

function wmGetTeacherAccountSheet_(ss) {
  if (!ss) return null;
  var names = ['7-1.교사관리_DB','7.설정센터_DB','7-2.교사학생_DB','교사관리_DB','교사_DB','Teacher_DB'];
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  return null;
}


/* WM_LMS_TEACHER_SUBLEADER_SCOPE_20260710_V2
 * TEACHER는 자기 담당 학생만, SUB_LEADER는 자기 소속 교사/반/학생만 조회합니다.
 * SUPER_ADMIN / SPECIAL_ADMIN의 기존 조회 범위는 변경하지 않습니다.
 */
function wmGetLmsRequestActor_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actorId = String(p.actorTeacherId || p.loginTeacherId || p.currentTeacherId || '').trim().toUpperCase();
  if (!actorId) return { found:false, teacherId:'', role:'' };
  return wmGetSettingCenterActor_({ parameter:{ actorTeacherId:actorId } });
}

function wmLmsCell_(row, names) {
  row = row || {};
  names = names || [];
  for (var i = 0; i < names.length; i++) {
    var key = String(names[i] || '').trim();
    if (key && row.hasOwnProperty(key) && String(row[key] || '').trim() !== '') return row[key];
  }
  return '';
}

function wmTeacherClassSet_(actor) {
  var set = {};
  wmNormalizeSettingClassList_(actor && actor.classes || '').forEach(function(v){ set[String(v).trim().toUpperCase()] = true; });
  return set;
}

/* 설정센터에서 한 교사에게 여러 SUB_LEADER가 선택될 수 있으므로 연결값을 목록으로 판정합니다. */
function wmSettingLeaderLinkSet_(value) {
  var set = {};
  String(value || '').split(/[,\n·|\/;]+/).forEach(function(v){
    var key = String(v || '').trim().toUpperCase();
    if (key && key !== '-' && key !== '추후선택') set[key] = true;
  });
  return set;
}

function wmSubLeaderScope_(actor) {
  var scope = { teacherIds:{}, teacherNames:{}, classes:{} };
  if (!actor || wmNormalizeLmsRole_(actor.role) !== 'SUB_LEADER') return scope;
  var actorId = String(actor.teacherId || '').trim().toUpperCase();
  var actorName = String(actor.name || '').trim().toUpperCase();
  try {
    var sheet = wmGetSettingCenterSheet_();
    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) return scope;
    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxId = wmFindHeaderIndex_(headers, ['교사ID']);
    var idxName = wmFindHeaderIndex_(headers, ['교사명','이름']);
    var idxLeader = wmFindHeaderIndex_(headers, ['담당리더','리더']);
    var idxClasses = wmFindHeaderIndex_(headers, ['담당반목록','담당반']);
    for (var r = 1; r < values.length; r++) {
      var rowId = idxId >= 0 ? String(values[r][idxId] || '').trim().toUpperCase() : '';
      var rowName = idxName >= 0 ? String(values[r][idxName] || '').trim().toUpperCase() : '';
      var rowLeaderLinks = wmSettingLeaderLinkSet_(idxLeader >= 0 ? values[r][idxLeader] : '');
      var owned = rowId === actorId || !!rowLeaderLinks[actorId] || !!(actorName && rowLeaderLinks[actorName]);
      if (!owned) continue;
      if (rowId) scope.teacherIds[rowId] = true;
      if (rowName) scope.teacherNames[rowName] = true;
      wmNormalizeSettingClassList_(idxClasses >= 0 ? values[r][idxClasses] : '').forEach(function(v){
        scope.classes[String(v || '').trim().toUpperCase()] = true;
      });
    }
  } catch (err) {}
  if (actorId) scope.teacherIds[actorId] = true;
  if (actorName) scope.teacherNames[actorName] = true;
  wmNormalizeSettingClassList_(actor.classes || '').forEach(function(v){ scope.classes[String(v).trim().toUpperCase()] = true; });
  return scope;
}

function wmStudentBelongsToActor_(row, actor) {
  if (!actor) return true;
  var role = wmNormalizeLmsRole_(actor.role);
  if (role !== 'TEACHER' && role !== 'SUB_LEADER') return true;
  var actorId = String(actor.teacherId || '').trim().toUpperCase();
  var actorName = String(actor.name || '').trim().toUpperCase();
  var teacherId = String(wmLmsCell_(row, ['교사ID','담당교사ID','현재담당교사ID','Teacher_ID','teacherId']) || '').trim().toUpperCase();
  var teacherName = String(wmLmsCell_(row, ['교사명','담당교사','현재담당교사','Teacher_Name','teacher']) || '').trim().toUpperCase();
  var leaderLinks = wmSettingLeaderLinkSet_(wmLmsCell_(row, ['담당리더','현재담당리더','리더','리더명','Leader','Leader_Name']) || '');
  var className = String(wmLmsCell_(row, ['반명','현재반','Class','반','담당반']) || '').trim().toUpperCase();
  if (role === 'TEACHER') {
    if (teacherId && teacherId === actorId) return true;
    if (teacherName && actorName && teacherName === actorName) return true;
    return !!(className && wmTeacherClassSet_(actor)[className]);
  }
  var scope = wmSubLeaderScope_(actor);
  if (leaderLinks[actorId] || (actorName && leaderLinks[actorName])) return true;
  if (teacherId && scope.teacherIds[teacherId]) return true;
  if (teacherName && scope.teacherNames[teacherName]) return true;
  return !!(className && scope.classes[className]);
}

function wmFilterStudentsForActor_(rows, actor) {
  rows = rows || [];
  var role = actor ? wmNormalizeLmsRole_(actor.role) : '';
  if (role !== 'TEACHER' && role !== 'SUB_LEADER') return rows;
  return rows.filter(function(row){ return wmStudentBelongsToActor_(row, actor); });
}

function wmRedactStudentSensitiveFieldsForActor_(rows, actor) {
  rows = rows || [];
  var role = actor ? wmNormalizeLmsRole_(actor.role) : '';
  if (role !== 'TEACHER' && role !== 'SUB_LEADER') return rows;
  var access = wmStudentDataAccessForActor_(actor);
  var parentFields = ['학부모명','학부모연락처','Parent_Name','Parent_Phone'];
  var paymentFields = ['최종결제','등록일','최초등록일','최근결제일','결제공백일수','재등록횟수','총등록기간'];
  return rows.map(function(row){
    var safe = {};
    Object.keys(row || {}).forEach(function(key){
      if (!access.parent && parentFields.indexOf(key) !== -1) return;
      if (!access.payment && paymentFields.indexOf(key) !== -1) return;
      safe[key] = row[key];
    });
    return safe;
  });
}

function wmStudentIdSet_(students) {
  var set = {};
  (students || []).forEach(function(row){
    var sid = String(wmLmsCell_(row, ['학생ID','studentId','Student_ID']) || '').trim().toUpperCase();
    if (sid) set[sid] = true;
  });
  return set;
}

function wmFilterRowsByStudentSet_(rows, studentSet, actor) {
  rows = rows || [];
  var role = actor ? wmNormalizeLmsRole_(actor.role) : '';
  if (role !== 'TEACHER' && role !== 'SUB_LEADER') return rows;
  return rows.filter(function(row){
    var sid = String(wmLmsCell_(row, ['학생ID','studentId','Student_ID']) || '').trim().toUpperCase();
    return !!(sid && studentSet[sid]);
  });
}

/* WM_SETTING_CENTER_DB_API_20260708_V1
 * 7-1.교사관리_DB / 7-2.반관리_DB 분리 관리 API입니다.
 * 교사/반 등록은 LMS 화면에서만 실행하고, 구글시트는 저장 결과 DB로만 사용합니다.
 */
function wmGetSettingCenterSheet_() {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('7-1.교사관리_DB');
  if (!sheet) sheet = ss.getSheetByName('7.설정센터_DB');
  if (!sheet) sheet = ss.getSheetByName('7-2.교사학생_DB');
  return sheet;
}

function wmSettingCenterHeaders_() {
  return ['교사명','교사ID','비밀번호','교사등급','담당리더','담당반목록','총담당반','총학생수','월학생수','월발송기록','상태','수정일','메모','학부모정보권한','결제정보권한'];
}

function wmNormalizeStudentDataAccessFlag_(value) {
  return /^(ON|Y|YES|TRUE|1|사용|공개)$/i.test(String(value || '').trim()) ? 'ON' : 'OFF';
}

function wmStudentDataAccessForActor_(actor) {
  var role = wmNormalizeLmsRole_(actor && actor.role || 'TEACHER');
  var adminLike = role === 'SUPER_ADMIN' || role === 'SPECIAL_ADMIN';
  return {
    parent:adminLike || wmNormalizeStudentDataAccessFlag_(actor && actor.parentAccess) === 'ON',
    payment:adminLike || wmNormalizeStudentDataAccessFlag_(actor && actor.paymentAccess) === 'ON'
  };
}

function wmEnsureSettingCenterHeader_(sheet) {
  var headers = wmSettingCenterHeaders_();
  var baseHeaders = headers.slice(0, 13);
  var accessHeaders = headers.slice(13);
  if (!sheet) throw new Error('7-1.교사관리_DB 시트를 찾을 수 없습니다.');
  var lastCol = Math.max(sheet.getLastColumn(), baseHeaders.length);
  var current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
  var changed = false;
  for (var i = 0; i < baseHeaders.length; i++) {
    if (current[i] !== baseHeaders[i]) {
      sheet.getRange(1, i + 1).setValue(baseHeaders[i]);
      current[i] = baseHeaders[i];
      changed = true;
    }
  }
  accessHeaders.forEach(function(header){
    if (current.indexOf(header) !== -1) return;
    sheet.getRange(1, current.length + 1).setValue(header);
    current.push(header);
    changed = true;
  });
  if (changed) SpreadsheetApp.flush();
  return current;
}

function wmNormalizeSettingClassList_(value) {
  var seen = {};
  var arr = String(value || '').split(/[,\n·|\/]+/).map(function(v){ return String(v || '').trim(); }).filter(function(v){ return v && v !== '-' && v !== '추후선택'; });
  var out = [];
  arr.forEach(function(v){ if (!seen[v]) { seen[v] = true; out.push(v); } });
  return out;
}

function wmCountStudentsByClassMap_() {
  var result = {};
  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('1.학생관리_DB');
    if (!sheet) return result;
    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) return result;
    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxClass = wmFindHeaderIndex_(headers, ['반명','Class','반','담당반','현재반']);
    var idxStatus = wmFindHeaderIndex_(headers, ['활성상태','상태']);
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var className = idxClass >= 0 ? String(row[idxClass] || '').trim() : '';
      if (!className) continue;
      var status = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : 'ACTIVE';
      if (/^(INACTIVE|퇴원|중지|비활성)$/i.test(status)) continue;
      result[className] = Number(result[className] || 0) + 1;
    }
  } catch (err) {}
  return result;
}

function wmBuildSettingCenterRowObject_(row, headers, studentCountMap) {
  var obj = {};
  headers.forEach(function(h, i){ obj[h] = String(row[i] || '').trim(); });
  obj['학부모정보권한'] = wmNormalizeStudentDataAccessFlag_(obj['학부모정보권한']);
  obj['결제정보권한'] = wmNormalizeStudentDataAccessFlag_(obj['결제정보권한']);
  var classList = wmNormalizeSettingClassList_(obj['담당반목록']);
  obj['담당반목록'] = classList.length ? classList.join(', ') : (obj['담당반목록'] || '-');
  obj['총담당반'] = classList.length ? String(classList.length) : (obj['총담당반'] || '-');
  var totalStudents = 0;
  classList.forEach(function(className){ totalStudents += Number((studentCountMap || {})[className] || 0); });
  obj['총학생수'] = classList.length ? String(totalStudents) : (obj['총학생수'] || '-');
  return obj;
}

/* WM_SETTING_CENTER_ROLE_DROPDOWN_V1
 * 설정센터 교사등급 드롭다운과 서버 반환 목록을 로그인 권한별로 제한합니다.
 * SUPER_ADMIN 4개 / SPECIAL_ADMIN 3개 / SUB_LEADER 2개 / TEACHER 0개
 */
function wmGetSettingCenterActor_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actorId = String(p.actorTeacherId || p.loginTeacherId || p.currentTeacherId || '').trim();
  var actorIdKey = actorId.toUpperCase();
  if (!actorId) return { found:false, teacherId:'', role:'TEACHER', name:'', leader:'', classes:'', parentAccess:'OFF', paymentAccess:'OFF' };

  var sheet = wmGetSettingCenterSheet_();
  var values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return { found:false, teacherId:actorId, role:'TEACHER', name:'', leader:'', classes:'', parentAccess:'OFF', paymentAccess:'OFF' };

  var headers = values[0].map(function(h){ return String(h || '').trim(); });
  var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
  var idxName = wmFindHeaderIndex_(headers, ['교사명','이름']);
  var idxRole = wmFindHeaderIndex_(headers, ['교사등급','권한']);
  var idxLeader = wmFindHeaderIndex_(headers, ['담당리더','리더']);
  var idxClasses = wmFindHeaderIndex_(headers, ['담당반목록','담당반']);
  var idxParentAccess = wmFindHeaderIndex_(headers, ['학부모정보권한','학부모관련권한']);
  var idxPaymentAccess = wmFindHeaderIndex_(headers, ['결제정보권한','결제관련권한']);

  for (var r = 1; r < values.length; r++) {
    var rowId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    if (rowId.toUpperCase() !== actorIdKey) continue;
    return {
      found:true,
      teacherId:rowId,
      name:idxName >= 0 ? String(values[r][idxName] || '').trim() : rowId,
      role:wmNormalizeLmsRole_(idxRole >= 0 ? values[r][idxRole] : 'TEACHER'),
      leader:idxLeader >= 0 ? String(values[r][idxLeader] || '').trim() : '',
      classes:idxClasses >= 0 ? String(values[r][idxClasses] || '').trim() : '',
      parentAccess:wmNormalizeStudentDataAccessFlag_(idxParentAccess >= 0 ? values[r][idxParentAccess] : 'OFF'),
      paymentAccess:wmNormalizeStudentDataAccessFlag_(idxPaymentAccess >= 0 ? values[r][idxPaymentAccess] : 'OFF')
    };
  }
  return { found:false, teacherId:actorId, role:'TEACHER', name:'', leader:'', classes:'', parentAccess:'OFF', paymentAccess:'OFF' };
}

function wmResolveSettingTeacherIdByName_(teacherName) {
  var target = String(teacherName || '').trim().toUpperCase();
  if (!target) return '';
  try {
    var sheet = wmGetSettingCenterSheet_();
    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) return '';
    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxName = wmFindHeaderIndex_(headers, ['교사명','이름']);
    var idxId = wmFindHeaderIndex_(headers, ['교사ID']);
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][idxName] || '').trim().toUpperCase() === target) return String(values[r][idxId] || '').trim();
    }
  } catch (err) {}
  return '';
}

function wmAllowedSettingGrades_(role) {
  var normalized = wmNormalizeLmsRole_(role);
  if (normalized === 'SUPER_ADMIN') return ['TEACHER','SUB_LEADER','SPECIAL_ADMIN','SUPER_ADMIN'];
  if (normalized === 'SPECIAL_ADMIN') return ['TEACHER','SUB_LEADER','SPECIAL_ADMIN'];
  /* SUB_LEADER는 설정센터 조회만 가능하며 교사/반 등록 권한은 없습니다. */
  if (normalized === 'SUB_LEADER') return [];
  return [];
}

function wmCanSeeSettingRow_(actor, rowObj) {
  if (!actor || !actor.found) return false;
  var role = wmNormalizeLmsRole_(actor.role);
  var rowRole = wmNormalizeLmsRole_(rowObj['교사등급'] || '');
  var rowTeacherId = String(rowObj['교사ID'] || '').trim().toUpperCase();
  var rowLeaderLinks = wmSettingLeaderLinkSet_(rowObj['담당리더'] || '');
  var actorId = String(actor.teacherId || '').trim().toUpperCase();
  var actorName = String(actor.name || '').trim().toUpperCase();

  if (role === 'SUPER_ADMIN') return true;
  if (role === 'SPECIAL_ADMIN') return rowRole !== 'SUPER_ADMIN';
  if (role === 'SUB_LEADER') {
    return rowTeacherId === actorId ||
      !!rowLeaderLinks[actorId] ||
      !!(actorName && rowLeaderLinks[actorName]);
  }
  return rowTeacherId === String(actor.teacherId || '').trim().toUpperCase();
}


/* WM_CLASS_MANAGEMENT_DB_READ_V1
 * 7-2.반관리_DB의 실제 11개 컬럼과 저장 행을 설정센터 응답에 포함합니다.
 */

/* WM_TEACHER_CLASS_FULL_SYNC_V1
 * 7-1.교사관리_DB의 담당반목록 전체를 7-2.반관리_DB에 동기화하고,
 * 1.학생관리_DB 반명 열의 데이터 유효성 목록을 전체 반 목록으로 갱신합니다.
 */
function wmSyncTeacherClassesToClassDb_() {
  var ss = getLmsSpreadsheet_();
  var teacherSheet = wmGetSettingCenterSheet_();
  var classSheet = ss.getSheetByName('7-2.반관리_DB');
  var studentSheet = ss.getSheetByName('1.학생관리_DB');

  if (!teacherSheet) return {success:false, message:'7-1.교사관리_DB 시트를 찾을 수 없습니다.'};
  if (!classSheet) return {success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'};
  if (!studentSheet) return {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};

  var officialClassHeaders = [
    '반명','담당교사ID','담당교사명','학생ID목록','학생이름목록',
    '학생수','상태','등록일','수정일','등록자','메모'
  ];

  var classLastColumn = classSheet.getLastColumn();
  if (classLastColumn < 1) return {success:false, message:'7-2.반관리_DB 헤더가 없습니다.'};
  var classHeaders = classSheet.getRange(1, 1, 1, classLastColumn).getDisplayValues()[0].map(function(h){
    return String(h || '').trim();
  });
  var missingClassHeaders = officialClassHeaders.filter(function(h){ return classHeaders.indexOf(h) === -1; });
  if (missingClassHeaders.length) {
    return {success:false, message:'7-2.반관리_DB 필수 컬럼 누락: ' + missingClassHeaders.join(', ')};
  }

  var teacherValues = teacherSheet.getDataRange().getDisplayValues();
  var teacherHeaders = teacherValues[0].map(function(h){ return String(h || '').trim(); });
  var idxTeacherName = wmFindHeaderIndex_(teacherHeaders, ['교사명','이름']);
  var idxTeacherId = wmFindHeaderIndex_(teacherHeaders, ['교사ID']);
  var idxTeacherClasses = wmFindHeaderIndex_(teacherHeaders, ['담당반목록','담당반']);
  if (idxTeacherName < 0 || idxTeacherId < 0 || idxTeacherClasses < 0) {
    return {success:false, message:'7-1.교사관리_DB의 교사명·교사ID·담당반목록 컬럼을 확인하세요.'};
  }

  var teacherByClass = {};
  for (var tr = 1; tr < teacherValues.length; tr++) {
    var teacherName = String(teacherValues[tr][idxTeacherName] || '').trim();
    var teacherId = String(teacherValues[tr][idxTeacherId] || '').trim();
    var classes = wmNormalizeSettingClassList_(teacherValues[tr][idxTeacherClasses]);
    classes.forEach(function(className){
      var normalizedClassName = String(className || '').trim();
      if (!normalizedClassName || normalizedClassName === '-') return;
      if (!teacherByClass[normalizedClassName]) {
        teacherByClass[normalizedClassName] = {
          teacherId:teacherId,
          teacherName:teacherName
        };
      }
    });
  }

  var studentValues = studentSheet.getDataRange().getDisplayValues();
  var studentHeaders = studentValues[0].map(function(h){ return String(h || '').trim(); });
  var idxStudentId = wmFindHeaderIndex_(studentHeaders, [
    '학생ID','Student_ID','studentId','student_id','StudentID','STUDENT_ID','학생 Id','학생id'
  ]);
  var idxStudentName = wmFindHeaderIndex_(studentHeaders, [
    '학생이름','학생명','이름','studentName','Student_Name','StudentName'
  ]);
  var idxStudentClass = wmFindHeaderIndex_(studentHeaders, ['반명','현재반','Class','반']);
  if (idxStudentId < 0 || idxStudentName < 0 || idxStudentClass < 0) {
    return {success:false, message:'1.학생관리_DB의 학생ID·학생이름·반명 컬럼을 확인하세요.'};
  }

  var studentsByClass = {};
  for (var sr = 1; sr < studentValues.length; sr++) {
    var className = String(studentValues[sr][idxStudentClass] || '').trim();
    if (!className) continue;
    if (!studentsByClass[className]) studentsByClass[className] = [];
    studentsByClass[className].push({
      id:String(studentValues[sr][idxStudentId] || '').trim(),
      name:String(studentValues[sr][idxStudentName] || '').trim()
    });
    if (!teacherByClass[className]) {
      teacherByClass[className] = {teacherId:'', teacherName:''};
    }
  }

  var allClassNames = Object.keys(teacherByClass).sort();
  var existingValues = classSheet.getDataRange().getValues();
  var idxClassNameDb = classHeaders.indexOf('반명');
  var existingRowByClass = {};
  for (var er = 1; er < existingValues.length; er++) {
    var existingClassName = String(existingValues[er][idxClassNameDb] || '').trim();
    if (existingClassName) existingRowByClass[existingClassName] = er + 1;
  }

  var now = new Date();
  var inserted = 0;
  var updated = 0;

  allClassNames.forEach(function(className){
    var teacher = teacherByClass[className] || {teacherId:'', teacherName:''};
    var students = studentsByClass[className] || [];
    var existingRow = existingRowByClass[className] || 0;

    var registeredAt = existingRow
      ? classSheet.getRange(existingRow, classHeaders.indexOf('등록일') + 1).getValue()
      : now;
    var status = existingRow
      ? String(classSheet.getRange(existingRow, classHeaders.indexOf('상태') + 1).getDisplayValue() || '').trim()
      : 'Active';
    var existingRegisteredBy = existingRow
      ? String(classSheet.getRange(existingRow, classHeaders.indexOf('등록자') + 1).getDisplayValue() || '').trim()
      : '';
    var existingMemo = existingRow
      ? String(classSheet.getRange(existingRow, classHeaders.indexOf('메모') + 1).getDisplayValue() || '').trim()
      : '';

    var rowObject = {
      '반명':className,
      '담당교사ID':teacher.teacherId,
      '담당교사명':teacher.teacherName,
      '학생ID목록':students.map(function(s){ return s.id; }).filter(Boolean).join(', '),
      '학생이름목록':students.map(function(s){ return s.name; }).filter(Boolean).join(', '),
      '학생수':students.length,
      '상태':status || 'Active',
      '등록일':registeredAt || now,
      '수정일':now,
      '등록자':existingRegisteredBy || '교사관리_DB 동기화',
      '메모':existingMemo || '7-1.교사관리_DB 담당반목록 자동동기화'
    };
    var outputRow = classHeaders.map(function(h){
      return Object.prototype.hasOwnProperty.call(rowObject, h) ? rowObject[h] : '';
    });

    var targetRow = existingRow || (classSheet.getLastRow() + 1);
    var targetRange = classSheet.getRange(targetRow, 1, 1, classHeaders.length);
    targetRange.clearDataValidations();
    targetRange.setValues([outputRow]);

    var statusIndex = classHeaders.indexOf('상태');
    if (statusIndex >= 0) {
      var statusValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Active','Inactive'], true)
        .setAllowInvalid(false)
        .build();
      classSheet.getRange(targetRow, statusIndex + 1).setDataValidation(statusValidation);
    }

    if (existingRow) updated++;
    else {
      inserted++;
      existingRowByClass[className] = targetRow;
    }
  });

  /* 학생관리_DB 반명 열 전체의 데이터 유효성 목록을 최신 전체 반 목록으로 교체합니다. */
  if (allClassNames.length && studentSheet.getMaxRows() >= 2) {
    var validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allClassNames, true)
      .setAllowInvalid(false)
      .build();
    var validationRange = studentSheet.getRange(
      2,
      idxStudentClass + 1,
      studentSheet.getMaxRows() - 1,
      1
    );
    validationRange.clearDataValidations();
    validationRange.setDataValidation(validationRule);
  }

  SpreadsheetApp.flush();

  return {
    success:true,
    classNames:allClassNames,
    inserted:inserted,
    updated:updated,
    total:allClassNames.length
  };
}

function wmGetClassManagementRowsForLms_() {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('7-2.반관리_DB');
  var studentNameById = {};
  try {
    var studentSheet = ss.getSheetByName('1.학생관리_DB');
    if (studentSheet && studentSheet.getLastRow() >= 2 && studentSheet.getLastColumn() >= 1) {
      var studentValues = studentSheet.getDataRange().getDisplayValues();
      var studentHeaders = studentValues[0].map(function(h){ return String(h || '').trim(); });
      var studentIdIndex = wmFindHeaderIndex_(studentHeaders, [
        '학생ID','Student_ID','studentId','student_id','StudentID','STUDENT_ID','학생 Id','학생id'
      ]);
      var studentNameIndex = wmFindHeaderIndex_(studentHeaders, [
        '학생이름','학생명','이름','studentName','Student_Name','StudentName'
      ]);
      if (studentIdIndex >= 0 && studentNameIndex >= 0) {
        for (var si = 1; si < studentValues.length; si++) {
          var lookupStudentId = String(studentValues[si][studentIdIndex] || '').trim();
          var lookupStudentName = String(studentValues[si][studentNameIndex] || '').trim();
          if (lookupStudentId) studentNameById[lookupStudentId.toUpperCase()] = lookupStudentName;
        }
      }
    }
  } catch (ignoreStudentNameLookup) {}

  var officialHeaders = [
    '반명','담당교사ID','담당교사명','학생ID목록','학생이름목록',
    '학생수','상태','등록일','수정일','등록자','메모'
  ];

  if (!sheet) {
    return {
      success:false,
      sheetName:'7-2.반관리_DB',
      headers:officialHeaders,
      rows:[],
      message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'
    };
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    return {
      success:true,
      sheetName:sheet.getName(),
      headers:officialHeaders,
      rows:[]
    };
  }

  var dataRange = sheet.getRange(1, 1, Math.max(lastRow, 1), lastColumn);
  var values = dataRange.getDisplayValues();
  var rawValues = dataRange.getValues();
  var sheetHeaders = values[0].map(function(h){ return String(h || '').trim(); });
  var headers = officialHeaders.filter(function(h){ return sheetHeaders.indexOf(h) !== -1; });
  if (!headers.length) headers = officialHeaders;

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var hasValue = values[r].some(function(v){ return String(v || '').trim(); });
    if (!hasValue) continue;
    var rowObj = {};
    headers.forEach(function(header){
      var idx = sheetHeaders.indexOf(header);
      rowObj[header] = idx >= 0 ? values[r][idx] : '';
    });

    /* 기존 행의 학생이름목록이 비어 있으면 학생ID목록을 기준으로 자동 복구하여 표시합니다. */
    if (!String(rowObj['학생이름목록'] || '').trim()) {
      var storedStudentIds = String(rowObj['학생ID목록'] || '')
        .split(',')
        .map(function(id){ return String(id || '').trim(); })
        .filter(function(id){ return !!id; });
      var recoveredNames = storedStudentIds.map(function(id){
        return studentNameById[id.toUpperCase()] || '';
      }).filter(function(name){ return !!name; });
      if (recoveredNames.length) rowObj['학생이름목록'] = recoveredNames.join(', ');
    }

    rowObj.__sheetRow = r + 1;
    var modifiedIndex = sheetHeaders.indexOf('수정일');
    var registeredIndex = sheetHeaders.indexOf('등록일');
    var modifiedRaw = modifiedIndex >= 0 ? rawValues[r][modifiedIndex] : '';
    var registeredRaw = registeredIndex >= 0 ? rawValues[r][registeredIndex] : '';
    var sortDate = modifiedRaw instanceof Date && !isNaN(modifiedRaw.getTime())
      ? modifiedRaw
      : (registeredRaw instanceof Date && !isNaN(registeredRaw.getTime()) ? registeredRaw : null);
    rowObj.__sortTime = sortDate ? sortDate.getTime() : 0;

    rows.push(rowObj);
  }

  rows.sort(function(a, b){
    var timeDiff = Number(b.__sortTime || 0) - Number(a.__sortTime || 0);
    if (timeDiff) return timeDiff;
    return Number(b.__sheetRow || 0) - Number(a.__sheetRow || 0);
  });

  return {
    success:true,
    sheetName:sheet.getName(),
    headers:headers,
    rows:rows
  };
}

/* WM_CLASS_DB_MISSING_CLASS_AUTO_SYNC_V1
 * 설정센터 진입 시 교사 담당반·학생 반명에 존재하지만 7-2.반관리_DB에 없는 반만 확인합니다.
 * 누락 반이 있을 때에만 기존 전체 동기화 함수를 1회 실행하여 하단에 모든 반이 표시되게 합니다.
 */
function wmEnsureAllClassesInClassDb_() {
  try {
    var ss = getLmsSpreadsheet_();
    var teacherSheet = wmGetSettingCenterSheet_();
    var studentSheet = ss.getSheetByName('1.학생관리_DB');
    var classSheet = ss.getSheetByName('7-2.반관리_DB');
    if (!teacherSheet || !studentSheet || !classSheet) return {success:false, skipped:true};

    var wanted = {};
    var teacherValues = teacherSheet.getDataRange().getDisplayValues();
    if (teacherValues.length) {
      var teacherHeaders = teacherValues[0].map(function(h){ return String(h || '').trim(); });
      var idxTeacherClasses = wmFindHeaderIndex_(teacherHeaders, ['담당반목록','담당반']);
      if (idxTeacherClasses >= 0) {
        for (var tr = 1; tr < teacherValues.length; tr++) {
          wmNormalizeSettingClassList_(teacherValues[tr][idxTeacherClasses]).forEach(function(name){
            name = String(name || '').trim();
            if (name && name !== '-') wanted[name] = true;
          });
        }
      }
    }

    var studentValues = studentSheet.getDataRange().getDisplayValues();
    if (studentValues.length) {
      var studentHeaders = studentValues[0].map(function(h){ return String(h || '').trim(); });
      var idxStudentClass = wmFindHeaderIndex_(studentHeaders, ['반명','현재반','Class','반']);
      if (idxStudentClass >= 0) {
        for (var sr = 1; sr < studentValues.length; sr++) {
          var studentClass = String(studentValues[sr][idxStudentClass] || '').trim();
          if (studentClass) wanted[studentClass] = true;
        }
      }
    }

    var existing = {};
    var classValues = classSheet.getDataRange().getDisplayValues();
    if (classValues.length) {
      var classHeaders = classValues[0].map(function(h){ return String(h || '').trim(); });
      var idxClass = wmFindHeaderIndex_(classHeaders, ['반명']);
      if (idxClass >= 0) {
        for (var cr = 1; cr < classValues.length; cr++) {
          var className = String(classValues[cr][idxClass] || '').trim();
          if (className) existing[className] = true;
        }
      }
    }

    var missing = Object.keys(wanted).filter(function(name){ return !existing[name]; });
    if (!missing.length) return {success:true, skipped:true, missing:[]};
    var syncResult = wmSyncTeacherClassesToClassDb_();
    syncResult.missingBeforeSync = missing;
    return syncResult;
  } catch (err) {
    return {success:false, error:String(err && err.message ? err.message : err)};
  }
}

function wmGetSettingCenterForLms_(e) {
  try {
    var actor = wmGetSettingCenterActor_(e);
    if (!actor.found) {
      return { success:false, message:'로그인 교사 정보를 확인할 수 없습니다.', rows:[], allowedGrades:[] };
    }

    /* 최초 조회 속도 보호: 반 전체 동기화는 저장 시 수행하며 조회 시에는 실행하지 않습니다. */
    var classSyncResult = {success:true, skipped:true, reason:'read_only_fast_path'};

    var sheet = wmGetSettingCenterSheet_();
    var headers = wmEnsureSettingCenterHeader_(sheet);
    var values = sheet.getDataRange().getDisplayValues();
    var studentCountMap = wmCountStudentsByClassMap_();
    var rows = [];

    for (var r = 1; r < values.length; r++) {
      var hasValue = values[r].some(function(v){ return String(v || '').trim(); });
      if (!hasValue) continue;
      var rowObj = wmBuildSettingCenterRowObject_(values[r], headers, studentCountMap);
      if (wmCanSeeSettingRow_(actor, rowObj)) rows.push(rowObj);
    }

    var classData = wmGetClassManagementRowsForLms_();
    var actorRole = wmNormalizeLmsRole_(actor.role);
    var allowedClassSet = {};
    if (actorRole === 'TEACHER') allowedClassSet = wmTeacherClassSet_(actor);
    if (actorRole === 'SUB_LEADER') allowedClassSet = wmSubLeaderScope_(actor).classes;
    if (actorRole === 'TEACHER' || actorRole === 'SUB_LEADER') {
      classData.rows = (classData.rows || []).filter(function(row){
        var className = String(wmLmsCell_(row, ['반명','Class','반']) || '').trim().toUpperCase();
        return !!(className && allowedClassSet[className]);
      });
    }

    return {
      success:true,
      sheetName:sheet.getName(),
      headers:headers,
      rows:rows,
      classSheetName:classData.sheetName,
      classHeaders:classData.headers || [],
      classRows:classData.rows || [],
      classReadSuccess:classData.success === true,
      classReadMessage:classData.message || '',
      classSyncResult:classSyncResult || {},
      actorRole:actor.role,
      actorParentAccess:wmNormalizeStudentDataAccessFlag_(actor.parentAccess),
      actorPaymentAccess:wmNormalizeStudentDataAccessFlag_(actor.paymentAccess),
      allowedGrades:wmAllowedSettingGrades_(actor.role),
      canRegister:wmAllowedSettingGrades_(actor.role).length > 0
    };
  } catch (err) {
    return { success:false, message:'7-1.교사관리_DB 조회 오류', error:String(err && err.message ? err.message : err), rows:[], allowedGrades:[] };
  }
}

function wmCheckSettingTeacherIdForLms_(e) {
  var teacherId = String((e && e.parameter && (e.parameter.teacherId || e.parameter.teacherID || e.parameter.id)) || '').trim();
  var teacherIdKey = teacherId.toLowerCase();
  if (!/^[A-Za-z0-9]{4,10}$/.test(teacherId)) return { success:false, available:false, message:'교사ID는 영어+숫자 4~10자리입니다.' };
  try {
    var sheet = wmGetSettingCenterSheet_();
    wmEnsureSettingCenterHeader_(sheet);
    var values = sheet.getDataRange().getDisplayValues();
    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][idxTeacherId] || '').trim().toLowerCase() === teacherIdKey) {
        return { success:true, available:false, message:'이미 사용 중인 교사ID입니다.' };
      }
    }
    return { success:true, available:true, message:'사용 가능한 교사ID입니다.' };
  } catch (err) {
    return { success:false, available:false, message:'교사ID 중복확인 오류', error:String(err && err.message ? err.message : err) };
  }
}

function wmSaveSettingTeacherForLms_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actor = wmGetSettingCenterActor_(e);
  if (!actor.found) return {success:false, message:'로그인 교사 정보를 확인할 수 없습니다.'};
  var teacherName = String(p.teacherName || p.name || '').trim();
  var teacherId = String(p.teacherId || p.teacherID || '').trim();
  var password = String(p.password || p.pw || '').trim();
  var grade = String(p.grade || p.role || 'TEACHER').trim();
  var leader = String(p.leader || p.담당리더 || '추후선택').trim();
  var className = String(p.className || p.class || p.담당반 || '추후선택').trim();
  var allowedGrades = wmAllowedSettingGrades_(actor.role);
  grade = wmNormalizeLmsRole_(grade);
  if (wmNormalizeLmsRole_(actor.role) === 'SUB_LEADER') leader = actor.name || actor.teacherId;
  if (allowedGrades.indexOf(grade) === -1) return {success:false, message:'현재 권한으로 등록할 수 없는 교사등급입니다.'};
  if (!teacherName) return {success:false, message:'교사명을 입력하세요.'};
  if (!/^[A-Za-z0-9]{4,10}$/.test(teacherId)) return {success:false, message:'교사ID는 영어+숫자 4~10자리입니다.'};
  if (!/^\d{4}$/.test(password)) return {success:false, message:'비밀번호는 숫자 4자리입니다.'};

  var lock = null;
  var locked = false;
  try {
    lock = LockService.getScriptLock();
    locked = lock.tryLock(2000);
    if (!locked) return {success:false, message:'교사등록이 동시에 처리 중입니다. 잠시 후 다시 저장하세요.'};

    var duplicate = wmCheckSettingTeacherIdForLms_({parameter:{teacherId:teacherId}});
    if (duplicate && duplicate.available === false) return duplicate;

    var sheet = wmGetSettingCenterSheet_();
    var headers = wmEnsureSettingCenterHeader_(sheet);
    var now = new Date();
    var classes = wmNormalizeSettingClassList_(className);
    var classList = classes.join(', ') || '-';
    var rowObj = {
      '교사명':teacherName, '교사ID':teacherId, '비밀번호':password, '교사등급':grade,
      '담당리더':leader || '추후선택', '담당반목록':classList,
      '총담당반':classes.length ? String(classes.length) : '-', '총학생수':'-', '월학생수':'-', '월발송기록':'-',
      '상태':'Active', '수정일':now, '메모':'', '학부모정보권한':'OFF', '결제정보권한':'OFF'
    };
    sheet.appendRow(headers.map(function(h){ return rowObj[h] || ''; }));

    var classSyncResults = [];
    classes.forEach(function(targetClass){
      try {
        classSyncResults.push(wmApplyClassTeacherChangeForLms_(getLmsSpreadsheet_(), targetClass, teacherId, teacherName, {updateClassRow:true}));
      } catch (syncErr) {
        classSyncResults.push({success:false, className:targetClass, message:String(syncErr && syncErr.message ? syncErr.message : syncErr)});
      }
    });
    var failedClassSync = classSyncResults.filter(function(result){ return result && result.success === false; });
    return {
      success:true,
      message:failedClassSync.length ? '교사등록 완료 (담당반 연결 ' + failedClassSync.length + '건 확인 필요)' : '교사등록 완료',
      teacherId:teacherId,
      classSyncResults:classSyncResults
    };
  } catch (err) {
    return {success:false, message:'교사등록 저장 오류', error:String(err && err.message ? err.message : err)};
  } finally {
    if (lock && locked) { try { lock.releaseLock(); } catch (ignore) {} }
  }
}

/* WM_SETTING_TEACHER_STATUS_UPDATE_20260710_V1
 * SUPER_ADMIN / SPECIAL_ADMIN만 설정센터의 기존 상태 컬럼을 Active / Inactive로 변경합니다.
 */
function wmSaveSettingTeacherStatusForLms_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actor = wmGetSettingCenterActor_(e);
  if (!actor.found) return { success:false, message:'로그인 교사 정보를 확인할 수 없습니다.' };

  var actorRole = wmNormalizeLmsRole_(actor.role);
  if (actorRole !== 'SUPER_ADMIN' && actorRole !== 'SPECIAL_ADMIN') {
    return { success:false, message:'교사 상태 변경 권한이 없습니다.' };
  }

  var teacherId = String(p.teacherId || p.teacherID || '').trim();
  var status = String(p.status || '').trim();
  if (!teacherId) return { success:false, message:'교사ID가 없습니다.' };
  if (!/^(Active|Inactive)$/i.test(status)) {
    return { success:false, message:'상태는 Active 또는 Inactive만 저장할 수 있습니다.' };
  }
  status = /^Inactive$/i.test(status) ? 'Inactive' : 'Active';

  try {
    var sheet = wmGetSettingCenterSheet_();
    var headers = wmEnsureSettingCenterHeader_(sheet);
    var values = sheet.getDataRange().getDisplayValues();
    var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
    var idxStatus = wmFindHeaderIndex_(headers, ['상태','교사상태','활성상태']);
    var idxUpdatedAt = wmFindHeaderIndex_(headers, ['수정일','수정일시']);

    if (idxTeacherId < 0 || idxStatus < 0) {
      return { success:false, message:'교사ID 또는 상태 컬럼을 찾을 수 없습니다.' };
    }

    var teacherIdKey = teacherId.toLowerCase();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][idxTeacherId] || '').trim().toLowerCase() !== teacherIdKey) continue;

      var rowObj = wmBuildSettingCenterRowObject_(values[r], headers, {});
      if (!wmCanSeeSettingRow_(actor, rowObj)) {
        return { success:false, message:'현재 권한으로 변경할 수 없는 교사입니다.' };
      }

      sheet.getRange(r + 1, idxStatus + 1).setValue(status);
      if (idxUpdatedAt >= 0) sheet.getRange(r + 1, idxUpdatedAt + 1).setValue(new Date());
      return { success:true, message:'교사 상태 저장 완료', teacherId:teacherId, status:status };
    }

    return { success:false, message:'해당 교사ID를 찾을 수 없습니다.' };
  } catch (err) {
    return { success:false, message:'교사 상태 저장 오류', error:String(err && err.message ? err.message : err) };
  }
}


/* WM_SETTING_TEACHER_DATA_ACCESS_TEST_V1
 * SUPER_ADMIN / SPECIAL_ADMIN이 SUB_LEADER / TEACHER의 학부모·결제정보 공개권한을 각각 저장합니다. */
function wmSaveSettingTeacherDataAccessForLms_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actor = wmGetSettingCenterActor_(e);
  if (!actor.found) return {success:false, message:'로그인 교사 정보를 확인할 수 없습니다.'};
  var actorRole = wmNormalizeLmsRole_(actor.role);
  if (actorRole !== 'SUPER_ADMIN' && actorRole !== 'SPECIAL_ADMIN') {
    return {success:false, message:'교사 정보권한 변경은 수퍼어드민 또는 스페셜어드민만 가능합니다.'};
  }

  var teacherId = String(p.teacherId || p.teacherID || '').trim();
  var parentRaw = String(p.parentAccess || p['학부모정보권한'] || '').trim().toUpperCase();
  var paymentRaw = String(p.paymentAccess || p['결제정보권한'] || '').trim().toUpperCase();
  if (!teacherId) return {success:false, message:'교사ID가 없습니다.'};
  if (['ON','OFF'].indexOf(parentRaw) === -1 || ['ON','OFF'].indexOf(paymentRaw) === -1) {
    return {success:false, message:'학부모정보권한과 결제정보권한은 ON 또는 OFF로 저장하세요.'};
  }

  try {
    var sheet = wmGetSettingCenterSheet_();
    var headers = wmEnsureSettingCenterHeader_(sheet);
    var values = sheet.getDataRange().getDisplayValues();
    var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
    var idxRole = wmFindHeaderIndex_(headers, ['교사등급','권한']);
    var idxParentAccess = wmFindHeaderIndex_(headers, ['학부모정보권한']);
    var idxPaymentAccess = wmFindHeaderIndex_(headers, ['결제정보권한']);
    var idxUpdatedAt = wmFindHeaderIndex_(headers, ['수정일','수정일시']);
    if (idxTeacherId < 0 || idxRole < 0 || idxParentAccess < 0 || idxPaymentAccess < 0) {
      return {success:false, message:'교사ID·교사등급·정보권한 컬럼을 확인하세요.'};
    }

    var teacherKey = teacherId.toLowerCase();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][idxTeacherId] || '').trim().toLowerCase() !== teacherKey) continue;
      var rowObj = wmBuildSettingCenterRowObject_(values[r], headers, {});
      if (!wmCanSeeSettingRow_(actor, rowObj)) return {success:false, message:'현재 권한으로 변경할 수 없는 교사입니다.'};
      var targetRole = wmNormalizeLmsRole_(values[r][idxRole]);
      if (targetRole !== 'TEACHER' && targetRole !== 'SUB_LEADER') {
        return {success:false, message:'학부모·결제정보 공개권한은 TEACHER와 SUB_LEADER에게만 설정합니다.'};
      }
      sheet.getRange(r + 1, idxParentAccess + 1).setValue(parentRaw);
      sheet.getRange(r + 1, idxPaymentAccess + 1).setValue(paymentRaw);
      if (idxUpdatedAt >= 0) sheet.getRange(r + 1, idxUpdatedAt + 1).setValue(new Date());
      SpreadsheetApp.flush();
      return {
        success:true,
        message:'교사 정보권한 저장 완료',
        teacherId:teacherId,
        parentAccess:parentRaw,
        paymentAccess:paymentRaw
      };
    }
    return {success:false, message:'해당 교사ID를 찾을 수 없습니다.'};
  } catch (err) {
    return {success:false, message:'교사 정보권한 저장 오류', error:String(err && err.message ? err.message : err)};
  }
}


/* WM_CLASS_STATUS_SAVE_V1
 * 7-2.반관리_DB의 선택 행 상태만 Active/Inactive로 저장합니다.
 * 반명 중복 가능성을 고려하여 화면에서 전달한 실제 시트 행번호를 우선 사용합니다.
 */
function wmSaveSettingClassStatusForLms_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actor = wmGetSettingCenterActor_(e);
  if (!actor.found) return {success:false, message:'로그인 교사 정보를 확인할 수 없습니다.'};
  var actorRole = wmNormalizeLmsRole_(actor.role);
  if (actorRole !== 'SUPER_ADMIN' && actorRole !== 'SPECIAL_ADMIN') return {success:false, message:'반 정보 변경은 수퍼어드민 또는 스페셜어드민만 가능합니다.'};

  var sheetRow = Number(p.sheetRow || p.row || 0);
  var className = String(p.className || p.class || '').trim();
  var requestedTeacherName = String(p.teacherName || '').trim();
  var requestedTeacherId = String(p.teacherId || '').trim();
  var status = /^Inactive$/i.test(String(p.status || '').trim()) ? 'Inactive' : 'Active';

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('7-2.반관리_DB');
    if (!sheet) return {success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'};
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) return {success:false, message:'저장할 반 데이터가 없습니다.'};

    var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
    var idxClassName = headers.indexOf('반명');
    var idxTeacherId = headers.indexOf('담당교사ID');
    var idxTeacherName = headers.indexOf('담당교사명');
    var idxStatus = headers.indexOf('상태');
    var idxUpdatedAt = headers.indexOf('수정일');
    if (idxClassName < 0 || idxTeacherId < 0 || idxTeacherName < 0 || idxStatus < 0) return {success:false, message:'반명·담당교사ID·담당교사명·상태 컬럼을 확인하세요.'};

    var targetRow = 0;
    if (sheetRow >= 2 && sheetRow <= lastRow) {
      var rowClassName = String(sheet.getRange(sheetRow, idxClassName + 1).getDisplayValue() || '').trim();
      if (!className || rowClassName === className) targetRow = sheetRow;
    }
    if (!targetRow && className) {
      var classNames = sheet.getRange(2, idxClassName + 1, lastRow - 1, 1).getDisplayValues();
      for (var r = classNames.length - 1; r >= 0; r--) {
        if (String(classNames[r][0] || '').trim() === className) { targetRow = r + 2; break; }
      }
    }
    if (!targetRow) return {success:false, message:'정보를 변경할 반 행을 찾을 수 없습니다.'};

    className = String(sheet.getRange(targetRow, idxClassName + 1).getDisplayValue() || '').trim();
    var currentTeacherName = String(sheet.getRange(targetRow, idxTeacherName + 1).getDisplayValue() || '').trim() || '추후선택';
    var currentTeacherId = String(sheet.getRange(targetRow, idxTeacherId + 1).getDisplayValue() || '').trim();
    var teacherName = requestedTeacherName || currentTeacherName;
    var teacherId = currentTeacherId;

    if (teacherName !== currentTeacherName || (requestedTeacherId && requestedTeacherId !== currentTeacherId)) {
      if (teacherName === '추후선택') {
        teacherId = '';
      } else {
        var teacherSheet = wmGetSettingCenterSheet_();
        var teacherHeaders = wmEnsureSettingCenterHeader_(teacherSheet);
        var teacherValues = teacherSheet.getDataRange().getDisplayValues();
        var idxSettingTeacherName = wmFindHeaderIndex_(teacherHeaders, ['교사명','이름']);
        var idxSettingTeacherId = wmFindHeaderIndex_(teacherHeaders, ['교사ID']);
        var foundTeacher = false;
        for (var tr = 1; tr < teacherValues.length; tr++) {
          var rowTeacherName = idxSettingTeacherName >= 0 ? String(teacherValues[tr][idxSettingTeacherName] || '').trim() : '';
          var rowTeacherId = idxSettingTeacherId >= 0 ? String(teacherValues[tr][idxSettingTeacherId] || '').trim() : '';
          var nameMatch = rowTeacherName === teacherName;
          var idMatch = requestedTeacherId && rowTeacherId.toUpperCase() === requestedTeacherId.toUpperCase();
          if (!nameMatch && !idMatch) continue;
          var teacherObj = wmBuildSettingCenterRowObject_(teacherValues[tr], teacherHeaders, {});
          if (!wmCanSeeSettingRow_(actor, teacherObj)) return {success:false, message:'현재 권한으로 연결할 수 없는 교사입니다.'};
          teacherName = rowTeacherName || teacherName;
          teacherId = rowTeacherId;
          foundTeacher = true;
          break;
        }
        if (!foundTeacher) return {success:false, message:'선택한 교사를 7-1.교사관리_DB에서 찾지 못했습니다.'};
        if (!teacherId) return {success:false, message:'선택한 교사의 교사ID가 없습니다.'};
      }
    }

    var teacherChanged = teacherName !== currentTeacherName || teacherId !== currentTeacherId;
    var statusCell = sheet.getRange(targetRow, idxStatus + 1);
    statusCell.clearDataValidations();
    statusCell.setValue(status);
    statusCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive'], true).setAllowInvalid(false).build());
    if (teacherChanged) {
      sheet.getRange(targetRow, idxTeacherName + 1).setValue(teacherName === '추후선택' ? '추후선택' : teacherName);
      var classTeacherIdCell = sheet.getRange(targetRow, idxTeacherId + 1);
      if (teacherId) {
        var classTeacherIdValidation = wmPrepareTeacherIdValidationForLms_(ss, sheet, [classTeacherIdCell.getA1Notation()]);
        if (!classTeacherIdValidation.success) return {success:false, message:classTeacherIdValidation.message || '담당교사ID 저장규칙 확인 실패'};
        classTeacherIdCell.setValue(teacherId);
      } else {
        classTeacherIdCell.clearDataValidations();
        classTeacherIdCell.clearContent();
      }
    }
    if (idxUpdatedAt >= 0) sheet.getRange(targetRow, idxUpdatedAt + 1).setValue(new Date());

    var teacherSync = {success:true, skipped:true};
    if (teacherChanged) {
      teacherSync = wmApplyClassTeacherChangeForLms_(ss, className, teacherId, teacherName, {
        updateClassRow:false,
        previousTeacherId:currentTeacherId,
        previousTeacherName:currentTeacherName
      });
      if (!teacherSync.success) return {success:false, partialSaved:true, message:'반 정보는 저장되었지만 학생·현재진행 담당교사 동기화에 실패했습니다.', error:teacherSync.message || ''};
    }

    SpreadsheetApp.flush();
    return {success:true, message:'반 정보 저장 완료', sheetRow:targetRow, className:className, teacherName:teacherName || '추후선택', teacherId:teacherId, status:status, teacherSync:teacherSync};
  } catch (err) {
    return {success:false, message:'반 정보 저장 오류', error:String(err && err.message ? err.message : err)};
  }
}

function wmSaveSettingClassForLms_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var actor = wmGetSettingCenterActor_(e);
  if (!actor.found) return { success:false, message:'로그인 교사 정보를 확인할 수 없습니다.' };

  var actorRole = wmNormalizeLmsRole_(actor.role);
  if (actorRole !== 'SUPER_ADMIN' && actorRole !== 'SPECIAL_ADMIN') {
    return { success:false, message:'반 생성과 학생 반 이동은 수퍼어드민 또는 스페셜어드민만 가능합니다.' };
  }

  var className = String(p.className || p.class || '').trim();
  var teacherName = String(p.teacherName || p.teacher || '').trim();
  var requestedTeacherId = String(p.teacherId || '').trim();
  var memo = String(p.memo || '').trim();
  var studentIds = String(p.studentIds || '').split(',').map(function(v){
    return String(v || '').trim();
  }).filter(function(v){ return !!v; });

  var studentRefs = [];
  try {
    var parsedRefs = JSON.parse(String(p.studentRefs || '[]'));
    if (Array.isArray(parsedRefs)) studentRefs = parsedRefs;
  } catch (ignoreRefs) {}

  if (!studentRefs.length) {
    studentRefs = studentIds.map(function(id){ return {id:id, name:''}; });
  }

  if (!className) return { success:false, message:'반명을 입력하세요.' };
  if (!teacherName) teacherName = '추후선택';
  if (!studentIds.length) return { success:false, message:'연결할 학생을 선택하세요.' };

  var lock = null;
  var lockAcquired = false;
  var stage = 'LOCK_INIT';

  try {
    lock = LockService.getScriptLock();
    stage = 'LOCK_WAIT';
    lockAcquired = lock.tryLock(2000);
    if (!lockAcquired) {
      return { success:false, message:'반등록이 동시에 처리 중입니다. 잠시 후 다시 저장하세요.', stage:stage };
    }

    stage = 'OPEN_LMS_SPREADSHEET';
    var ss = getLmsSpreadsheet_();
    var now = new Date();

    /* 1. 저장 전에 7-2.반관리_DB 구조를 먼저 검증하여 중간 저장을 방지합니다. */
    stage = 'VALIDATE_CLASS_DB';
    var classSheet = ss.getSheetByName('7-2.반관리_DB');
    if (!classSheet) return { success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.', stage:stage };

    var classLastColumn = classSheet.getLastColumn();
    if (classLastColumn < 1) return { success:false, message:'7-2.반관리_DB의 헤더가 없습니다.', stage:stage };

    var classHeaders = classSheet.getRange(1, 1, 1, classLastColumn).getDisplayValues()[0].map(function(h){
      return String(h || '').trim();
    });
    var requiredClassHeaders = [
      '반명','담당교사ID','담당교사명','학생ID목록','학생이름목록',
      '학생수','상태','등록일','수정일','등록자','메모'
    ];
    var missingClassHeaders = requiredClassHeaders.filter(function(h){ return classHeaders.indexOf(h) === -1; });
    if (missingClassHeaders.length) {
      return {
        success:false,
        message:'7-2.반관리_DB 필수 컬럼을 찾을 수 없습니다: ' + missingClassHeaders.join(', '),
        stage:stage
      };
    }

    /* 2. 담당교사는 7-1.교사관리_DB에서 이름과 ID를 확정합니다. */
    stage = 'RESOLVE_TEACHER';
    var settingSheet = wmGetSettingCenterSheet_();
    var headers = wmEnsureSettingCenterHeader_(settingSheet);
    var values = settingSheet.getDataRange().getDisplayValues();
    var idxTeacherName = wmFindHeaderIndex_(headers, ['교사명','이름']);
    var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
    var idxClasses = wmFindHeaderIndex_(headers, ['담당반목록','담당반']);
    var idxTotalClass = wmFindHeaderIndex_(headers, ['총담당반']);
    var idxUpdatedAt = wmFindHeaderIndex_(headers, ['수정일','수정일시']);
    var idxMemo = wmFindHeaderIndex_(headers, ['메모','비고']);
    if (idxTeacherName < 0 || idxTeacherId < 0 || idxClasses < 0) {
      return { success:false, message:'7-1.교사관리_DB의 교사명·교사ID·담당반목록 컬럼을 확인하세요.', stage:stage };
    }

    var teacherRowIndex = -1;
    var selectedTeacherId = '';
    var teacherUnassigned = teacherName === '추후선택';
    if (!teacherUnassigned) {
      for (var r = 1; r < values.length; r++) {
        var rowTeacherName = idxTeacherName >= 0 ? String(values[r][idxTeacherName] || '').trim() : '';
        var rowTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
        var nameMatch = rowTeacherName === teacherName;
        var idMatch = requestedTeacherId && rowTeacherId.toUpperCase() === requestedTeacherId.toUpperCase();
        if (!nameMatch && !idMatch) continue;

        var candidateObj = wmBuildSettingCenterRowObject_(values[r], headers, {});
        if (!wmCanSeeSettingRow_(actor, candidateObj)) {
          return { success:false, message:'현재 권한으로 연결할 수 없는 교사입니다.', stage:stage };
        }
        teacherRowIndex = r;
        teacherName = rowTeacherName || teacherName;
        selectedTeacherId = rowTeacherId;
        break;
      }
      if (teacherRowIndex < 0) return { success:false, message:'선택한 교사를 7-1.교사관리_DB에서 찾지 못했습니다.', stage:stage };
      if (!selectedTeacherId) return { success:false, message:'선택한 교사의 교사ID가 없습니다.', stage:stage };
    }

    /* 3. 학생은 1.학생관리_DB에서 ID와 이름을 확정하고 반·교사를 연결합니다. */
    stage = 'RESOLVE_AND_UPDATE_STUDENTS';
    var studentSheet = ss.getSheetByName('1.학생관리_DB');
    if (!studentSheet) return { success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.', stage:stage };

    var studentValues = studentSheet.getDataRange().getDisplayValues();
    var studentHeaders = studentValues[0].map(function(h){ return String(h || '').trim(); });
    var idxStudentId = wmFindHeaderIndex_(studentHeaders, ['학생ID','Student_ID','studentId','student_id','StudentID','STUDENT_ID','학생 Id','학생id']);
    var idxStudentName = wmFindHeaderIndex_(studentHeaders, ['학생이름','학생명','이름','studentName','Student_Name','StudentName']);
    var idxClassName = wmFindHeaderIndex_(studentHeaders, ['반명','현재반','Class','반']);
    var idxStudentTeacherName = wmFindHeaderIndex_(studentHeaders, ['교사명','현재담당교사','담당교사']);
    var idxStudentTeacherId = wmFindHeaderIndex_(studentHeaders, ['교사ID','현재담당교사ID','담당교사ID']);
    var idxStudentMemo = wmFindHeaderIndex_(studentHeaders, ['비고','메모']);

    if (idxStudentId < 0 || idxStudentName < 0 || idxClassName < 0) {
      return { success:false, message:'학생관리_DB의 학생ID·학생이름·반명 컬럼을 확인하세요.', stage:stage };
    }

    var wantedById = {};
    studentRefs.forEach(function(ref){
      var id = String(ref && ref.id || '').trim();
      if (id) wantedById[id.toUpperCase()] = {id:id, name:String(ref && ref.name || '').trim()};
    });
    studentIds.forEach(function(id){
      var key = String(id || '').trim().toUpperCase();
      if (key && !wantedById[key]) wantedById[key] = {id:String(id).trim(), name:''};
    });

    var resolvedStudents = [];
    var movedCount = 0;
    var affectedClassNames = [className];
    var classCellA1 = [];
    var teacherNameCellA1 = [];
    var teacherIdCellA1 = [];
    var memoUpdates = [];
    for (var sr = 1; sr < studentValues.length; sr++) {
      var sid = String(studentValues[sr][idxStudentId] || '').trim();
      var sidKey = sid.toUpperCase();
      if (!sidKey || !wantedById[sidKey]) continue;

      var studentName =
        String(studentValues[sr][idxStudentName] || '').trim() ||
        String(wantedById[sidKey].name || '').trim() ||
        sid;
      var previousClass = String(studentValues[sr][idxClassName] || '').trim();
      if (previousClass && previousClass !== className) {
        movedCount++;
        affectedClassNames.push(previousClass);
      }

      classCellA1.push(studentSheet.getRange(sr + 1, idxClassName + 1).getA1Notation());
      studentValues[sr][idxClassName] = className;
      if (idxStudentTeacherName >= 0) studentValues[sr][idxStudentTeacherName] = teacherUnassigned ? '' : teacherName;
      if (idxStudentTeacherId >= 0) studentValues[sr][idxStudentTeacherId] = teacherUnassigned ? '' : selectedTeacherId;
      if (idxStudentTeacherName >= 0) {
        teacherNameCellA1.push(studentSheet.getRange(sr + 1, idxStudentTeacherName + 1).getA1Notation());
      }
      if (idxStudentTeacherId >= 0) {
        teacherIdCellA1.push(studentSheet.getRange(sr + 1, idxStudentTeacherId + 1).getA1Notation());
      }
      if (idxStudentMemo >= 0 && previousClass !== className) {
        var oldMemo = String(studentValues[sr][idxStudentMemo] || '').trim();
        var moveText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
          ' 반이동(' + (previousClass || '미배정') + '→' + className + ', 처리자:' + (actor.name || actor.teacherId) + ')';
        memoUpdates.push({
          row:sr + 1,
          value:oldMemo ? oldMemo + ' | ' + moveText : moveText
        });
      }

      resolvedStudents.push({id:sid, name:studentName});
      delete wantedById[sidKey];
    }

    var missingStudentIds = Object.keys(wantedById).map(function(key){ return wantedById[key].id; });
    if (!resolvedStudents.length) {
      return { success:false, message:'선택한 학생을 학생관리_DB에서 찾지 못했습니다: ' + missingStudentIds.join(', '), stage:stage };
    }

    /* 선택 학생의 반·교사 공통값을 RangeList로 일괄 저장합니다. */
    stage = 'BATCH_UPDATE_STUDENTS';
    if (classCellA1.length) {
      var classRanges = studentSheet.getRangeList(classCellA1);
      classRanges.clearDataValidations();
      classRanges.setValue(className);
    }
    if (teacherNameCellA1.length) {
      studentSheet.getRangeList(teacherNameCellA1).setValue(teacherUnassigned ? '' : teacherName);
    }
    if (teacherIdCellA1.length) {
      var teacherIdRanges = studentSheet.getRangeList(teacherIdCellA1);
      teacherIdRanges.clearDataValidations();
      teacherIdRanges.setValue(teacherUnassigned ? '' : selectedTeacherId);
    }
    memoUpdates.forEach(function(update){
      studentSheet.getRange(update.row, idxStudentMemo + 1).setValue(update.value);
    });
    /* 현재진행_DB 동기화는 반관리_DB 저장과 담당교사 확정 뒤 한 번만 배치 실행합니다. */
    var currentProgressStudentInfoSync = {success:true, skipped:true};

    /* 4. 7-1.교사관리_DB의 담당반목록을 갱신합니다. */
    stage = 'UPDATE_TEACHER_CLASS';
    if (!teacherUnassigned) {
      var classes = wmNormalizeSettingClassList_(values[teacherRowIndex][idxClasses]);
      if (classes.indexOf(className) === -1) classes.push(className);
      settingSheet.getRange(teacherRowIndex + 1, idxClasses + 1).setValue(classes.join(', '));
      if (idxTotalClass >= 0) settingSheet.getRange(teacherRowIndex + 1, idxTotalClass + 1).setValue(classes.length);
      if (idxUpdatedAt >= 0) settingSheet.getRange(teacherRowIndex + 1, idxUpdatedAt + 1).setValue(now);
      if (idxMemo >= 0) settingSheet.getRange(teacherRowIndex + 1, idxMemo + 1).setValue('반등록: ' + className);
    }

    /* 5. 11개 출처를 한 행으로 매핑하고 반명 기준으로 신규등록 또는 수정합니다. */
    stage = 'UPSERT_CLASS_DB';
    var classValues = classSheet.getDataRange().getValues();
    var idxClassDbName = classHeaders.indexOf('반명');
    var existingRowIndex = -1;
    for (var cr = 1; cr < classValues.length; cr++) {
      if (String(classValues[cr][idxClassDbName] || '').trim() === className) {
        existingRowIndex = cr + 1;
        break;
      }
    }

    var existingRegisteredAt = '';
    var existingStatus = '';
    var previousClassTeacherId = '';
    var previousClassTeacherName = '';
    if (existingRowIndex > 0) {
      existingRegisteredAt = classSheet.getRange(existingRowIndex, classHeaders.indexOf('등록일') + 1).getValue();
      existingStatus = String(classSheet.getRange(existingRowIndex, classHeaders.indexOf('상태') + 1).getDisplayValue() || '').trim();
      previousClassTeacherId = String(classSheet.getRange(existingRowIndex, classHeaders.indexOf('담당교사ID') + 1).getDisplayValue() || '').trim();
      previousClassTeacherName = String(classSheet.getRange(existingRowIndex, classHeaders.indexOf('담당교사명') + 1).getDisplayValue() || '').trim();
    }

    var classRow = {
      '반명': className,
      '담당교사ID': selectedTeacherId,
      '담당교사명': teacherName,
      '학생ID목록': resolvedStudents.map(function(s){ return s.id; }).join(', '),
      '학생이름목록': resolvedStudents.map(function(s){ return s.name; }).join(', '),
      '학생수': resolvedStudents.length,
      '상태': existingStatus || 'Active',
      '등록일': existingRegisteredAt || now,
      '수정일': now,
      '등록자': actor.name || actor.teacherId || '',
      '메모': memo || '반등록'
    };

    var outputRow = classHeaders.map(function(h){
      return Object.prototype.hasOwnProperty.call(classRow, h) ? classRow[h] : '';
    });
    var targetClassRow = existingRowIndex > 0 ? existingRowIndex : classSheet.getLastRow() + 1;
    var targetClassRange = classSheet.getRange(targetClassRow, 1, 1, classHeaders.length);

    /* 다른 DB에서 복사된 잘못된 드롭다운 규칙이 신규 행에 남아 저장을 막지 않도록
       대상 행의 유효성 규칙만 제거하고, 상태 셀에 필요한 규칙만 다시 설정합니다. */
    targetClassRange.clearDataValidations();
    targetClassRange.setValues([outputRow]);

    var statusColumnIndex = classHeaders.indexOf('상태');
    if (statusColumnIndex >= 0) {
      var statusValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Active','Inactive'], true)
        .setAllowInvalid(false)
        .build();
      classSheet.getRange(targetClassRow, statusColumnIndex + 1).setDataValidation(statusValidation);
    }

    SpreadsheetApp.flush();
    stage = 'REBUILD_CLASS_ROSTERS';
    var rosterSync = wmRebuildClassRostersFromStudentDb_(ss, affectedClassNames);
    if (!rosterSync.success) return {success:false, message:'반 명단 재구성에 실패했습니다.', error:rosterSync.message || '', stage:stage};

    stage = 'SYNC_CLASS_TEACHER';
    var teacherSync = wmApplyClassTeacherChangeForLms_(ss, className, selectedTeacherId, teacherName, {
      updateClassRow:false,
      previousTeacherId:previousClassTeacherId,
      previousTeacherName:previousClassTeacherName
    });
    if (!teacherSync.success) return {success:false, partialSaved:true, message:'반 저장은 완료되었지만 담당교사 동기화에 실패했습니다.', error:teacherSync.message || '', stage:stage};
    currentProgressStudentInfoSync = teacherSync.currentProgressStudentInfoSync || {success:true, skipped:true};

    stage = 'VERIFY_CLASS_DB';
    var verifyRow = targetClassRow;
    var savedValues = classSheet.getRange(verifyRow, 1, 1, classHeaders.length).getDisplayValues()[0];
    var savedObject = {};
    classHeaders.forEach(function(h, i){ savedObject[h] = savedValues[i]; });
    if (String(savedObject['반명'] || '').trim() !== className) {
      return { success:false, message:'반관리_DB 저장 후 반명 확인에 실패했습니다.', stage:stage };
    }

    return {
      success:true,
      message:(existingRowIndex > 0 ? '반정보 수정 완료' : '반등록 완료') + ': 학생 ' + resolvedStudents.length + '명 연결' + (missingStudentIds.length ? ' / ' + missingStudentIds.length + '명 제외' : ''),
      className:className,
      linkedCount:resolvedStudents.length,
      movedCount:movedCount,
      savedStudentIds:resolvedStudents.map(function(student){ return student.id; }),
      skippedStudentIds:missingStudentIds,
      classRow:savedObject,
      rosterSync:rosterSync,
      teacherSync:teacherSync,
      currentProgressStudentInfoSync:currentProgressStudentInfoSync,
      stage:'COMPLETE'
    };
  } catch (err) {
    return {
      success:false,
      message:'반등록 저장 오류 (' + stage + ')',
      error:String(err && err.message ? err.message : err),
      stage:stage
    };
  } finally {
    if (lock && lockAcquired) {
      try { lock.releaseLock(); } catch (ignore) {}
    }
  }
}

function wmFindHeaderIndex_(headers, names) {
  headers = headers || [];
  for (var i = 0; i < names.length; i++) {
    var target = String(names[i] || '').trim().toLowerCase();
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j] || '').trim().toLowerCase() === target) return j;
    }
  }
  return -1;
}

function wmNormalizeLmsRole_(role) {
  var text = String(role || '').trim().toUpperCase();
  var compact = text.replace(/[\s_-]+/g, '');
  if (text === '슈퍼관리자' || text === '최고관리자' || text === 'SUPER' || text === 'SUPER_ADMIN' || text === 'SUPER ADMIN' || compact === 'SUPERADMIN') return 'SUPER_ADMIN';
  if (text === '관리자' || text === '특별관리자' || text === 'SPECIAL_ADMIN' || text === 'SPECIAL ADMIN' || text === 'ADMIN' || compact === 'SPECIALADMIN') return 'SPECIAL_ADMIN';
  if (text === '리더' || text === '담당리더' || text === 'SUB_LEADER' || text === 'SUB LEADER' || text === 'LEADER' || compact === 'SUBLEADER') return 'SUB_LEADER';
  return 'TEACHER';
}



function parseMapStudentPayload_(payloadText) {
  var text = String(payloadText || '').trim();
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (err1) {}

  try {
    var decoded = Utilities.newBlob(Utilities.base64Decode(text)).getDataAsString('UTF-8');
    return JSON.parse(decoded);
  } catch (err2) {}

  try {
    var webText = text.replace(/-/g, '+').replace(/_/g, '/');
    while (webText.length % 4) {
      webText += '=';
    }
    var decodedWeb = Utilities.newBlob(Utilities.base64Decode(webText)).getDataAsString('UTF-8');
    return JSON.parse(decodedWeb);
  } catch (err3) {}

  return null;
}

function escapeHtmlForMap_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function extractLevelNumberForMap_(learningAssign) {
  var match = String(learningAssign || '').match(/(\d+)/);
  return match ? Number(match[1]) : 0;
}


function normalizeCurrentSetForMap_(setId, learningAssign) {
  /* WM_CURRENT_SET_NORMALIZE_20260617_V1
   * 학생관리_DB 현재세트에 2004-1-4처럼 비정상 접두값이 들어와도
   * 화면/학습맵 계산에는 4-1-4 형태의 실제 세트ID만 사용합니다.
   * DB 원본값은 여기서 수정하지 않고 응답값만 정규화합니다.
   */
  var text = String(setId || '').trim().toUpperCase().replace(/^WM/, '');
  var match = text.match(/^(\d+)-(\d+)-(\d+)$/);

  if (match) {
    var level = Number(match[1]);
    var part = Number(match[2]);
    var set = Number(match[3]);

    if (level >= 3 && level <= 13) {
      return level + '-' + part + '-' + set;
    }

    var tail2 = level % 100;
    var tail1 = level % 10;
    if (tail2 >= 3 && tail2 <= 13) {
      return tail2 + '-' + part + '-' + set;
    }
    if (tail1 >= 3 && tail1 <= 9) {
      return tail1 + '-' + part + '-' + set;
    }
  }

  return '';
}

function getStudentBasicInfoForMap_(studentId) {
  /* WM_STUDENT_PROFILE_FOR_SAVE_V1
   * 1.학생관리_DB를 학생ID 기준으로 조회해서
   * Map 상단 표시 + Study 저장 시 학습기록_DB 기본정보 보강에 함께 사용합니다.
   */
  var result = {
    studentId: String(studentId || '').trim(),
    studentName: '',
    school: '',
    grade: '',
    className: '',
    teacherName: '',
    learningAssign: '',
    currentSet: '',
    learningMode: buildDefaultLearningMode_()
  };

  if (!result.studentId) {
    return result;
  }

  var cacheKey = wmCacheKey_('STUDENT_PROFILE', result.studentId);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && cached.studentId) {
    /* WM_STUDENT_PROFILE_EMPTY_CACHE_BYPASS_20260704_V1
     * 이전 조회 실패로 빈 학생정보가 캐시되면 저장 payload의 학교/학년/반명/교사명이
     * 계속 공백으로 들어갑니다. 핵심 정보가 모두 비어 있으면 DB를 다시 읽습니다.
     */
    var cachedHasCompleteProfile = !!(
      String(cached.studentName || '').trim() &&
      String(cached.school || '').trim() &&
      String(cached.grade || '').trim() &&
      String(cached.className || '').trim() &&
      String(cached.teacherName || '').trim() &&
      String(cached.learningAssign || '').trim()
    );
    if (cachedHasCompleteProfile) {
      cached.studentId = result.studentId; return cached;
    }
  }

  try {
    var ss = getLmsSpreadsheet_();
    if (!ss) {
      return result;
    }

    var sheet = ss.getSheetByName('1.학생관리_DB');
    if (!sheet) {
      return result;
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1 || sheet.getLastRow() < 2) {
      return result;
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });

    var idxStudentId = headers.indexOf('학생ID');
    var idxStudentName = headers.indexOf('학생이름');
    var idxSchool = headers.indexOf('학교');
    var idxGrade = headers.indexOf('학년');
    var idxClassId = findFirstHeaderIndex_(headers, ['Class', '반ID', '반', '클래스ID']);
    var idxClassName = findFirstHeaderIndex_(headers, ['반명', '학급명', '클래스명', '수업반명', '반이름']);
    var idxTeacherName = headers.indexOf('교사명');
    if (idxTeacherName === -1) idxTeacherName = headers.indexOf('교사');
    if (idxTeacherName === -1) idxTeacherName = headers.indexOf('담당교사');
    var idxLearningAssign = findFirstHeaderIndex_(headers, ['최초배정', '학습배정']);
    var idxCurrentSet = headers.indexOf('현재세트');
    if (idxLearningAssign === -1) idxLearningAssign = headers.indexOf('학습레벨');
    if (idxLearningAssign === -1) idxLearningAssign = headers.indexOf('레벨');

    if (idxStudentId === -1) {
      return result;
    }

    var studentRows = wmGetStudentRowNumbersFast_(sheet, headers, result.studentId);
    if (studentRows.length) {
        var row = sheet.getRange(studentRows[0], 1, 1, lastCol).getDisplayValues()[0];
        result.studentId = String(row[idxStudentId] || result.studentId).trim(); result.studentName = idxStudentName >= 0 ? String(row[idxStudentName] || '').trim() : '';
        result.school = idxSchool >= 0 ? String(row[idxSchool] || '').trim() : '';
        result.grade = idxGrade >= 0 ? String(row[idxGrade] || '').trim() : '';
        var rawClassId = idxClassId >= 0 ? String(row[idxClassId] || '').trim() : '';
        var rawClassName = idxClassName >= 0 ? String(row[idxClassName] || '').trim() : '';
        result.className = resolveClassDisplayName_(rawClassId, rawClassName);
        result.teacherName = idxTeacherName >= 0 ? String(row[idxTeacherName] || '').trim() : '';
        result.learningAssign = idxLearningAssign >= 0 ? String(row[idxLearningAssign] || '').trim() : '';
        result.currentSet = normalizeCurrentSetForMap_(idxCurrentSet >= 0 ? String(row[idxCurrentSet] || '').trim() : '', result.learningAssign || '4레벨');
        result.learningMode = buildLearningModeFromRow_(row, headers);

        if (/^[0-9]+$/.test(result.learningAssign)) {
          result.learningAssign = result.learningAssign + '레벨';
        }

        wmCachePutJson_(cacheKey, result, 300);
        return result;
    }
  } catch (err) {}

  wmCachePutJson_(cacheKey, result, 120);
  return result;
}


function findFirstHeaderIndex_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx >= 0) {
      return idx;
    }
  }
  return -1;
}

function normalizeLearningAssignForLms_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var m = text.match(/^(3|4|5|6|7|8|9|10|11|12|13)/);
  if (!m) return text;
  return m[1] + '레벨';
}

function normalizeS4SilhouetteStep_(value) {
  var text = String(value || '').trim();
  if (!text) return '4단계';
  if (text === '0' || text === '0단계' || text === '전체' || text === 'FULL' || text === 'full') return '1단계';
  if (text === '1' || text.indexOf('1단계') === 0) return '1단계';
  if (text === '2' || text.indexOf('2단계') === 0) return '2단계';
  if (text === '3' || text.indexOf('3단계') === 0) return '3단계';
  if (text === '4' || text.indexOf('4단계') === 0) return '4단계';
  return '4단계';
}

function normalizeS6TestMode_(mode) {
  var text = String(mode || '').trim();
  if (!text) return '기본전체';
  if (text === '기본전체' || text === 'FULL' || text === 'full') return '기본전체';
  if (text === '객관식' || text === '객관식2' || text === '객관식(2개)' || text === 'MC_ONLY') return '객관식(2개)';
  if (text === '객과식+영한') return '객관식+영한';
  if (text === '객관식+영한' || text === 'MC_ENKO') return '객관식+영한';
  if (text === '객관식+한영' || text === 'MC_KOEN') return '객관식+한영';
  return '기본전체';
}

function normalizeLevelCompleteCondition_(value) {
  var text = String(value || '').trim();
  if (!text) return '2회';
  if (text === '1' || text === '1회') return '1회';
  if (text === '2' || text === '2회') return '2회';
  if (text === '3' || text === '3회') return '3회';
  if (text === '추가' || text === '추가회차') return '추가';
  return '2회';
}

function normalizeLevelExtraRounds_(value) {
  var text = String(value || '').trim();
  if (!text) return 0;
  var n = Number(String(text).replace(/[^0-9.-]/g, ''));
  if (!isFinite(n) || n < 0) return 0;
  return Math.min(3, Math.floor(n));
}


/* WM_LEVEL_TARGET_ROUND_ENGINE_20260616_V1
 * 2단계: 레벨완료조건 + 레벨회차추가를 하나의 목표회차로 계산하는 공통 엔진입니다.
 * 현재 단계에서는 저장/맵/학습 흐름을 바꾸지 않고, 이후 순차완주/레벨완료/다음레벨 자동열림에서 공통으로 호출합니다.
 */
function getLevelBaseRoundsFromCondition_(condition) {
  var normalized = normalizeLevelCompleteCondition_(condition);

  if (normalized === '1회') return 1;
  if (normalized === '2회') return 2;
  if (normalized === '3회') return 3;

  /* '추가'는 별도 기본회차가 아니라 추가회차 운용 상태로 봅니다.
   * 기존 기본값은 2회이므로, 별도 기준이 없을 때는 2회를 기준으로 계산합니다.
   */
  if (normalized === '추가') return 2;

  return 2;
}

function buildLevelTargetRoundGoal_(levelCompleteCondition, levelExtraRounds) {
  /* WM_LEVEL_COMPLETE_CONDITION_DB_ONLY_20260616_V4
   * 최종 기준: 공식 목표회차는 1.학생관리_DB의 '레벨완료조건' 하나만 사용합니다.
   * 레벨회차추가 컬럼값은 더 이상 목표회차/색상/다음레벨 자동열림에 반영하지 않습니다.
   * - 1회: 공식 1회 순차완주
   * - 2회: 공식 2회 순차완주
   * - 3회: 공식 3회 순차완주
   * - 추가: 별도 운영 선택값으로 유지하되, 현재 엔진에서는 기존 기본 기준인 2회로 계산합니다.
   */
  var condition = normalizeLevelCompleteCondition_(levelCompleteCondition);
  var baseRounds = getLevelBaseRoundsFromCondition_(condition);
  var extraRounds = 0;
  var targetRounds = baseRounds;

  if (targetRounds < 1) targetRounds = 1;
  if (targetRounds > 3) targetRounds = 3;

  return {
    레벨완료조건: condition,
    기본회차: baseRounds,
    레벨회차추가: extraRounds,
    목표회차: targetRounds,
    isExtraApplied: condition === '추가',
    applyScope: 'NEXT_SET',
    기준: '학생관리_DB.레벨완료조건'
  };
}

function buildLevelTargetRoundGoalFromLearningMode_(learningMode) {
  var mode = learningMode || buildDefaultLearningMode_();
  return buildLevelTargetRoundGoal_(mode.레벨완료조건, mode.레벨회차추가);
}

function getStudentLevelTargetRoundGoal_(studentId) {
  var mode = getStudentLearningMode_(studentId, '');
  return buildLevelTargetRoundGoalFromLearningMode_(mode);
}


/* WM_LEVEL_SEQUENTIAL_ROUND_ENGINE_20260616_V1
 * 3단계: 레벨 전체를 순차 완주했는지 판정하는 공통 엔진입니다.
 * 아직 학습맵/UI/저장 흐름에 직접 연결하지 않고, 이후 레벨완료/자동열림에서 공통으로 호출합니다.
 * 원칙: N회차 공식 인정은 레벨 안의 모든 세트가 순서대로 N회 이상 완료되었을 때만 인정합니다.
 */
function normalizeLevelPlainSetId_(setId) {
  var text = String(setId || '').trim().toUpperCase();
  text = text.replace(/^WM/, '');
  var match = text.match(/^(\d+)-(\d+)-(\d+)$/);
  if (!match) return '';
  return Number(match[1]) + '-' + Number(match[2]) + '-' + Number(match[3]);
}

/* WM_LEVEL_SET_STRUCTURE_FINAL_20260616_V1
 * 3~13레벨의 실제 마지막 세트를 기준으로 세트 수를 고정합니다.
 * 3: 3-3-2 / 4: 4-3-2 / 5: 5-4-2 / 6: 6-4-2
 * 7~9: x-4-10 / 10~13: x-5-10
 */
function getLearningMapLevelPartSetCounts_(level) {
  /* WM_STUDENT_DB_MODE_APPLY_20260616_V1
   * 학생관리_DB 학습배정 기준과 동일한 실제 세트 구조입니다.
   * 3: 3-3-2 / 4: 4-3-2 / 5: 5-3-3 / 6: 6-3-3
   * 7~9: x-4-10 / 10~13: x-5-10
   */
  var n = Number(level || 0);

  if (n === 3) return [10, 10, 2];
  if (n === 4) return [10, 10, 2];
  if (n === 5) return [10, 10, 3];
  if (n === 6) return [10, 10, 3];

  if (n >= 7 && n <= 9) return [10, 10, 10, 10];
  if (n >= 10 && n <= 13) return [10, 10, 10, 10, 10];

  return [];
}

function getLearningMapPartSetCount_(level, part) {
  var counts = getLearningMapLevelPartSetCounts_(level);
  var idx = Number(part || 0) - 1;
  if (idx < 0 || idx >= counts.length) return 0;
  return Number(counts[idx] || 0);
}

function isValidLearningMapSetId_(setId) {
  var parts = parseLearningMapPlainSetId_(setId);
  if (!parts) return false;
  var maxSet = getLearningMapPartSetCount_(parts.level, parts.part);
  return maxSet > 0 && parts.set >= 1 && parts.set <= maxSet;
}

function getLearningMapLastSetIdByLevel_(level) {
  var n = Number(level || 0);
  var counts = getLearningMapLevelPartSetCounts_(n);
  if (!counts.length) return '';
  return n + '-' + counts.length + '-' + counts[counts.length - 1];
}


function getLevelSetSequence_(level) {
  var levelNum = Number(level || 0);
  var partCounts = getLearningMapLevelPartSetCounts_(levelNum);
  var result = [];

  if (!levelNum || !partCounts.length) return result;

  for (var part = 1; part <= partCounts.length; part++) {
    var maxSet = Number(partCounts[part - 1] || 0);
    for (var set = 1; set <= maxSet; set++) {
      result.push(levelNum + '-' + part + '-' + set);
    }
  }

  return result;
}

function buildRepeatCountMapFromSetStatusMap_(setStatusMap) {
  var repeatMap = {};
  if (!setStatusMap || typeof setStatusMap !== 'object') return repeatMap;

  Object.keys(setStatusMap).forEach(function(key) {
    var item = setStatusMap[key] || {};
    var setId = normalizeLevelPlainSetId_(item.setId || key);
    if (!setId) return;

    var repeatCount = Number(item.repeatCount || 0);
    if (!isFinite(repeatCount) || repeatCount < 0) repeatCount = 0;

    repeatMap[setId] = Math.max(Number(repeatMap[setId] || 0), Math.floor(repeatCount));
  });

  return repeatMap;
}

function countSequentialCompletedRoundsForLevel_(level, repeatMap, maxRounds) {
  var sequence = getLevelSetSequence_(level);
  var limit = Number(maxRounds || 0);
  var completedRounds = 0;

  if (!sequence.length || !repeatMap || typeof repeatMap !== 'object') return 0;
  if (!limit || limit < 1) limit = 6;

  for (var round = 1; round <= limit; round++) {
    var fullRoundComplete = true;

    for (var i = 0; i < sequence.length; i++) {
      if (Number(repeatMap[sequence[i]] || 0) < round) {
        fullRoundComplete = false;
        break;
      }
    }

    if (!fullRoundComplete) break;
    completedRounds = round;
  }

  return completedRounds;
}

function getNextSequentialSetForLevelRound_(level, repeatMap, targetRounds) {
  var sequence = getLevelSetSequence_(level);
  var target = Number(targetRounds || 1);
  if (target < 1) target = 1;

  for (var round = 1; round <= target; round++) {
    for (var i = 0; i < sequence.length; i++) {
      var setId = sequence[i];
      if (Number((repeatMap || {})[setId] || 0) < round) {
        return {
          level: Number(level || 0),
          officialRound: round,
          setId: setId,
          setIndex: i + 1,
          totalSets: sequence.length,
          levelComplete: false
        };
      }
    }
  }

  return {
    level: Number(level || 0),
    officialRound: target,
    setId: '',
    setIndex: sequence.length,
    totalSets: sequence.length,
    levelComplete: true
  };
}

function buildSequentialLevelRoundState_(level, setStatusMap, targetRoundGoal) {
  var goal = targetRoundGoal || buildLevelTargetRoundGoal_('', 0);
  var targetRounds = Number(goal.목표회차 || goal.targetRounds || 1);
  var repeatMap = buildRepeatCountMapFromSetStatusMap_(setStatusMap);
  var completedRounds = countSequentialCompletedRoundsForLevel_(level, repeatMap, targetRounds);
  var next = getNextSequentialSetForLevelRound_(level, repeatMap, targetRounds);

  return {
    level: Number(level || 0),
    목표회차: targetRounds,
    공식완료회차: completedRounds,
    현재공식회차: next.levelComplete ? targetRounds : next.officialRound,
    다음공식세트: next.setId,
    현재회차세트순번: next.setIndex,
    전체세트수: next.totalSets,
    레벨완료: !!next.levelComplete,
    repeatCountMap: repeatMap
  };
}



/* WM_LEVEL_COMPLETION_ENGINE_20260616_V1
 * 4단계: 목표회차와 순차완주 결과를 비교하여 레벨완료 여부만 판정하는 공통 엔진입니다.
 * 아직 다음레벨 자동열림/UI/저장 흐름에는 직접 연결하지 않습니다.
 */
function buildLevelCompletionState_(level, setStatusMap, learningMode) {
  var goal = buildLevelTargetRoundGoalFromLearningMode_(learningMode || buildDefaultLearningMode_());
  var sequential = buildSequentialLevelRoundState_(level, setStatusMap, goal);
  var targetRounds = Number(goal.목표회차 || 1);
  var completedRounds = Number(sequential.공식완료회차 || 0);
  var isComplete = completedRounds >= targetRounds;

  return {
    level: Number(level || 0),
    레벨완료: isComplete,
    목표회차: targetRounds,
    공식완료회차: completedRounds,
    다음공식세트: isComplete ? '' : String(sequential.다음공식세트 || ''),
    현재공식회차: isComplete ? targetRounds : Number(sequential.현재공식회차 || 1),
    현재회차세트순번: Number(sequential.현재회차세트순번 || 0),
    전체세트수: Number(sequential.전체세트수 || 0),
    goal: goal,
    sequential: sequential
  };
}

function isLevelCompleteBySequentialRounds_(level, setStatusMap, learningMode) {
  return !!buildLevelCompletionState_(level, setStatusMap, learningMode).레벨완료;
}



/* WM_NEXT_LEVEL_UNLOCK_ENGINE_20260616_V1
 * 5단계: 레벨완료 판정 결과를 기준으로 다음 레벨 자동열림 대상 세트를 계산하는 공통 엔진입니다.
 * 아직 학습맵/UI/저장 흐름에는 직접 연결하지 않고, 다음 단계에서 이 결과를 맵 표시와 현재세트 갱신에 연결합니다.
 */
function getNextLevelFirstSetId_(level) {
  var levelNum = Number(level || 0);
  if (!levelNum || levelNum < 3) return '';
  if (levelNum >= 13) return '';
  return String(levelNum + 1) + '-1-1';
}

function buildNextLevelUnlockState_(level, setStatusMap, learningMode) {
  var completion = buildLevelCompletionState_(level, setStatusMap, learningMode || buildDefaultLearningMode_());
  var currentLevel = Number(level || completion.level || 0);
  var nextSetId = completion.레벨완료 ? getNextLevelFirstSetId_(currentLevel) : '';

  return {
    level: currentLevel,
    레벨완료: !!completion.레벨완료,
    다음레벨자동열림: !!nextSetId,
    다음레벨: nextSetId ? currentLevel + 1 : '',
    다음레벨첫세트: nextSetId,
    목표회차: Number(completion.목표회차 || 0),
    공식완료회차: Number(completion.공식완료회차 || 0),
    completion: completion
  };
}

function getAutoUnlockedSetAfterLevelComplete_(level, setStatusMap, learningMode) {
  return String(buildNextLevelUnlockState_(level, setStatusMap, learningMode).다음레벨첫세트 || '');
}

function normalizeLearningMode_(mode) {
  /* WM_LEARNING_MODE_ENGINE_V1
   * 기존 단일 학습모드 문자열을 S6테스트모드로 안전하게 흡수합니다.
   */
  return normalizeS6TestMode_(mode);
}

function buildDefaultLearningMode_() {
  return {
    S4실루엣단계: '전체',
    S6테스트모드: '기본전체',
    레벨완료조건: '2회',
    레벨회차추가: 0,
    목표회차: 2,
    applyScope: 'NEXT_SET',
    note: '학습모드 변경은 현재 진행 중 세트가 아니라 다음 새 세트부터 적용됩니다.'
  };
}

function buildLearningModeFromRow_(row, headers) {
  var mode = buildDefaultLearningMode_();
  if (!row || !headers) return mode;

  var idxS4 = headers.indexOf('S4실루엣단계');
  var idxS6 = headers.indexOf('S6테스트모드');
  var idxComplete = headers.indexOf('레벨완료조건');
  var idxExtra = headers.indexOf('레벨회차추가');
  var idxLegacyMode = headers.indexOf('학습모드');

  mode.S4실루엣단계 = normalizeS4SilhouetteStep_(idxS4 >= 0 ? row[idxS4] : '');
  mode.S6테스트모드 = normalizeS6TestMode_(idxS6 >= 0 ? row[idxS6] : (idxLegacyMode >= 0 ? row[idxLegacyMode] : ''));
  mode.레벨완료조건 = normalizeLevelCompleteCondition_(idxComplete >= 0 ? row[idxComplete] : '');
  mode.레벨회차추가 = normalizeLevelExtraRounds_(idxExtra >= 0 ? row[idxExtra] : '');
  mode.목표회차 = buildLevelTargetRoundGoal_(mode.레벨완료조건, mode.레벨회차추가).목표회차;

  return mode;
}

function getStudentLearningMode_(studentId, className) {
  studentId = String(studentId || '').trim().toUpperCase();
  var result = buildDefaultLearningMode_();

  if (!studentId) return result;

  var cacheKey = wmCacheKey_('STUDENT_MODE', studentId);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached) return cached;

  try {
    var profile = getStudentBasicInfoForMap_(studentId);
    if (profile && profile.learningMode) {
      wmCachePutJson_(cacheKey, profile.learningMode, 300);
      return profile.learningMode;
    }
  } catch (profileErr) {}

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('1.학생관리_DB');
    if (!sheet) return result;

    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) return result;

    var headers = values[0].map(function(h) { return String(h || '').trim(); });
    var idxStudentId = headers.indexOf('학생ID');
    if (idxStudentId === -1) return result;

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowStudentId = String(row[idxStudentId] || '').trim().toUpperCase();
      if (rowStudentId === studentId) {
        result = buildLearningModeFromRow_(row, headers);
        wmCachePutJson_(cacheKey, result, 300);
        return result;
      }
    }

    wmCachePutJson_(cacheKey, result, 120);
    return result;
  } catch (err) {
    return result;
  }
}


/* WM_STUDENT_BASIC_INFO_API_20260704_V1
 * 1.학생관리_DB 단독 검사용 공개 API입니다.
 * 기존 내부 함수 getStudentBasicInfoForMap_는 수정하지 않고,
 * 학생ID 조회 결과를 Map/Study/저장엔진에서 쓰는 동일 구조로 반환합니다.
 * 사용 예: ?action=getStudentBasicInfoForMap&studentId=S4821
 */
function getStudentBasicInfoForMapApi_(e) {
  var studentId = '';
  if (e && e.parameter) {
    studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || e.parameter.wmStudentId || '').trim();
  }

  if (!studentId) {
    return outputResult(e, {
      success: false,
      message: '학생ID가 없습니다.',
      studentId: ''
    });
  }

  var profile = getStudentBasicInfoForMap_(studentId);
  var found = !!String(profile.studentName || profile.school || profile.grade || profile.className || profile.teacherName || profile.learningAssign || profile.currentSet || '').trim();

  return outputResult(e, {
    success: found,
    message: found ? '학생관리_DB 조회 성공' : '학생관리_DB에서 학생ID를 찾지 못했습니다.',
    studentId: profile.studentId || studentId,
    studentName: profile.studentName || '',
    school: profile.school || '',
    grade: profile.grade || '',
    className: profile.className || '',
    teacherName: profile.teacherName || '',
    learningAssign: profile.learningAssign || '',
    currentSet: profile.currentSet || '',
    learningMode: profile.learningMode || buildDefaultLearningMode_(),
    raw: {
      '학생ID': profile.studentId || studentId,
      '학생이름': profile.studentName || '',
      '학교': profile.school || '',
      '학년': profile.grade || '',
      'Class': profile.className || '',
      '교사명': profile.teacherName || '',
      '학습배정': profile.learningAssign || '',
      '현재세트': profile.currentSet || '',
      '학습모드': profile.learningMode || buildDefaultLearningMode_()
    }
  });
}

function getStudentLearningModeApi_(e) {
  var studentId = '';
  var requestedSetId = '';
  if (e && e.parameter) {
    studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || '').trim();
    requestedSetId = String(e.parameter.set_id || e.parameter.setId || e.parameter.Set_ID || '').trim().toUpperCase();
  }

  var profile = getStudentBasicInfoForMap_(studentId);
  var learningMode = getStudentLearningMode_(studentId, '');
  var modeSource = '1.학생관리_DB_NEXT_SET';
  var progressMode = wmGetCurrentProgressCacheForStudent_(studentId);
  var progressSetId = String(progressMode && progressMode['Set_ID'] || '').trim().toUpperCase();
  if (requestedSetId && progressSetId && normalizeStudySetIdForCompare_(requestedSetId) === normalizeStudySetIdForCompare_(progressSetId)) {
    learningMode = Object.assign({}, learningMode, {
      S4실루엣단계: normalizeS4SilhouetteStep_(progressMode['S4실루엣단계'] || learningMode.S4실루엣단계),
      S6테스트모드: normalizeS6TestMode_(progressMode['S6테스트모드'] || learningMode.S6테스트모드)
    });
    modeSource = '8.현재진행_DB_CURRENT_SET_LOCK';
  }

  return outputResult(e, {
    success: true,
    studentId: studentId,
    learningAssign: profile.learningAssign || '',
    currentSet: buildInitialSetIdFromLearningAssign(profile.learningAssign || ''),
    learningMode: learningMode,
    modeSource: modeSource,
    modeColumns: {
      학습배정: profile.learningAssign || '',
      S4실루엣단계: learningMode.S4실루엣단계,
      S6테스트모드: learningMode.S6테스트모드,
      레벨완료조건: learningMode.레벨완료조건,
      레벨회차추가: learningMode.레벨회차추가
    }
  });
}

function wmNextLevelFromCurrentSetForLms_(currentSet) {
  var match = String(currentSet || '').trim().toUpperCase().match(/^WM(\d+)-\d+-\d+$/);
  return match ? Number(match[1]) + 1 : '';
}

function wmNormalizeStudentLearningDate_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  var date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  if (date.getFullYear() !== Number(match[1]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) return null;
  return text;
}

function wmStudentLearningDateValue_(isoDate) {
  var parts = String(isoDate || '').split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

/* WM_LEARNING_DATE_SHEET_SOURCE_V2
 * 1.학생관리_DB의 빈 학습시작일은 같은 행의 등록일로 실제 저장합니다.
 * 사용자가 입력한 학습시작일·학습종료일은 덮어쓰지 않습니다.
 * 6.성적등급_DB는 미발행 행만 학생관리_DB 날짜와 동기화합니다.
 * 반영완료 행의 발행 당시 기간은 이후 어떤 동기화에서도 변경하지 않습니다.
 */
function wmSyncLearningDatesToSheets_() {
  var ss = getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  var gradeSheet = ss.getSheetByName('6.성적등급_DB');
  if (!studentSheet) return {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};

  var studentLastRow = studentSheet.getLastRow();
  var studentLastCol = studentSheet.getLastColumn();
  if (studentLastRow < 2 || studentLastCol < 1) return {success:true, studentUpdated:0, gradeUpdated:0};
  var studentHeaders = studentSheet.getRange(1, 1, 1, studentLastCol).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
  var studentIdIndex = studentHeaders.indexOf('학생ID');
  var studentStartIndex = studentHeaders.indexOf('학습시작일');
  var studentEndIndex = studentHeaders.indexOf('학습종료일');
  var studentRegisterIndex = studentHeaders.indexOf('등록일');
  if (studentIdIndex < 0 || studentStartIndex < 0 || studentEndIndex < 0 || studentRegisterIndex < 0) {
    return {success:false, message:'1.학생관리_DB의 학생ID·학습시작일·학습종료일·등록일 컬럼을 확인하세요.'};
  }

  var studentRowCount = studentLastRow - 1;
  var studentValues = studentSheet.getRange(2, 1, studentRowCount, studentLastCol).getValues();
  var studentStartValues = studentValues.map(function(row){ return [row[studentStartIndex]]; });
  var studentPeriodMap = {};
  var studentUpdated = 0;
  for (var s = 0; s < studentValues.length; s++) {
    var studentId = String(studentValues[s][studentIdIndex] || '').trim().toUpperCase();
    if (!studentId) continue;
    if (!String(studentValues[s][studentStartIndex] || '').trim() && String(studentValues[s][studentRegisterIndex] || '').trim()) {
      studentValues[s][studentStartIndex] = studentValues[s][studentRegisterIndex];
      studentStartValues[s][0] = studentValues[s][studentRegisterIndex];
      studentUpdated++;
    }
    studentPeriodMap[studentId] = {
      start:studentValues[s][studentStartIndex] || '',
      end:studentValues[s][studentEndIndex] || ''
    };
  }
  if (studentUpdated) {
    studentSheet.getRange(2, studentStartIndex + 1, studentRowCount, 1)
      .setValues(studentStartValues)
      .setNumberFormat('yyyy-MM-dd');
  }

  if (!gradeSheet || gradeSheet.getLastRow() < 2) {
    SpreadsheetApp.flush();
    return {success:true, studentUpdated:studentUpdated, gradeUpdated:0};
  }
  var gradeLastRow = gradeSheet.getLastRow();
  var gradeLastCol = gradeSheet.getLastColumn();
  var gradeHeaders = gradeSheet.getRange(1, 1, 1, gradeLastCol).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
  var gradeIdIndex = gradeHeaders.indexOf('학생ID');
  var gradeStartIndex = gradeHeaders.indexOf('학습시작일');
  var gradeEndIndex = gradeHeaders.indexOf('학습종료일');
  var gradeStatusIndex = gradeHeaders.indexOf('성적표반영상태');
  if (gradeIdIndex < 0 || gradeStartIndex < 0 || gradeEndIndex < 0 || gradeStatusIndex < 0) {
    return {success:false, message:'6.성적등급_DB의 학생ID·학습시작일·학습종료일·성적표반영상태 컬럼을 확인하세요.'};
  }

  var gradeRowCount = gradeLastRow - 1;
  var gradeValues = gradeSheet.getRange(2, 1, gradeRowCount, gradeLastCol).getValues();
  var gradeStartValues = gradeValues.map(function(row){ return [row[gradeStartIndex]]; });
  var gradeEndValues = gradeValues.map(function(row){ return [row[gradeEndIndex]]; });
  var gradeUpdated = 0;
  for (var g = 0; g < gradeValues.length; g++) {
    var gradeStatus = String(gradeValues[g][gradeStatusIndex] || '').trim();
    if (gradeStatus === '반영완료') continue;
    var gradeStudentId = String(gradeValues[g][gradeIdIndex] || '').trim().toUpperCase();
    var period = studentPeriodMap[gradeStudentId];
    if (!period) continue;
    var oldStart = String(gradeValues[g][gradeStartIndex] || '').trim();
    var oldEnd = String(gradeValues[g][gradeEndIndex] || '').trim();
    var newStart = period.start || '';
    var newEnd = period.end || '';
    if (oldStart !== String(newStart || '').trim() || oldEnd !== String(newEnd || '').trim()) gradeUpdated++;
    gradeStartValues[g][0] = newStart;
    gradeEndValues[g][0] = newEnd;
  }
  if (gradeUpdated) {
    gradeSheet.getRange(2, gradeStartIndex + 1, gradeRowCount, 1).setValues(gradeStartValues).setNumberFormat('yyyy-MM-dd');
    gradeSheet.getRange(2, gradeEndIndex + 1, gradeRowCount, 1).setValues(gradeEndValues).setNumberFormat('yyyy-MM-dd');
  }
  SpreadsheetApp.flush();
  return {success:true, studentUpdated:studentUpdated, gradeUpdated:gradeUpdated};
}

/* WM_ADMIN_TEACHER_ID_VALIDATION_SYNC_V1_20260821
 * 학생/반 저장에서 교사ID 데이터확인 규칙이 과거 목록(T001,T002...)에 고정되어
 * 정상 교사ID 저장까지 막지 않도록, 실제 7-1.교사관리_DB의 교사ID 목록으로
 * 저장 대상 셀의 규칙만 갱신합니다. 다른 학생/학습 데이터는 변경하지 않습니다. */
function wmPrepareTeacherIdValidationForLms_(ss, targetSheet, a1List) {
  ss = ss || getLmsSpreadsheet_();
  targetSheet = targetSheet || null;
  a1List = (a1List || []).filter(function(a1){ return !!String(a1 || '').trim(); });
  if (!targetSheet || !a1List.length) return {success:true, skipped:true};

  var teacherSheet = ss.getSheetByName('7-1.교사관리_DB');
  if (!teacherSheet) return {success:false, message:'7-1.교사관리_DB 시트를 찾을 수 없습니다.'};
  var values = teacherSheet.getDataRange().getDisplayValues();
  if (!values || !values.length) return {success:false, message:'7-1.교사관리_DB 데이터가 없습니다.'};
  var headers = values[0].map(function(v){ return String(v || '').trim(); });
  var idxTeacherId = wmFindHeaderIndex_(headers, ['교사ID']);
  if (idxTeacherId < 0) return {success:false, message:'7-1.교사관리_DB 교사ID 컬럼을 찾을 수 없습니다.'};

  var ids = [];
  var seen = {};
  for (var r = 1; r < values.length; r++) {
    var id = String(values[r][idxTeacherId] || '').trim();
    var key = id.toUpperCase();
    if (!id || seen[key]) continue;
    seen[key] = true;
    ids.push(id);
  }
  if (!ids.length) return {success:false, message:'7-1.교사관리_DB에 등록된 교사ID가 없습니다.'};

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ids, true)
    .setAllowInvalid(false)
    .build();
  var ranges = targetSheet.getRangeList(a1List).getRanges();
  ranges.forEach(function(range){
    range.clearDataValidations();
    range.setDataValidation(rule);
  });
  return {success:true, teacherCount:ids.length, targetCount:ranges.length};
}

function saveStudentLearningModeApi_(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    function hasParam_(names) {
      for (var hp = 0; hp < names.length; hp++) {
        if (Object.prototype.hasOwnProperty.call(p, names[hp])) return true;
      }
      return false;
    }

    var studentId = String(p.studentId || p.studentID || p.sid || '').trim().toUpperCase();
    if (!studentId) return outputResult(e, {success:false, message:'학생ID가 없습니다.'});

    var activeOnlyAllowed = {
      mode:true,
      action:true,
      actorTeacherId:true,
      studentId:true,
      studentID:true,
      sid:true,
      activeStatus:true,
      '활성상태':true,
      ts:true,
      callback:true
    };

    var activeOnlyRequest =
      hasParam_(['activeStatus','활성상태']) &&
      Object.keys(p).every(function(key) {
        return !!activeOnlyAllowed[key];
      });

    if (activeOnlyRequest) {
      return outputResult(e, wmSaveStudentActiveStatusFast_(e));
    }

    var actor = wmGetLmsRequestActor_(e);
    if (String(p.actorTeacherId || '').trim() && !actor.found) {
      return outputResult(e, {success:false, message:'로그인 교사 정보를 확인할 수 없습니다.'});
    }

    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('1.학생관리_DB');
    if (!sheet) return outputResult(e, {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'});

    var values = sheet.getDataRange().getDisplayValues();
    if (!values || !values.length) return outputResult(e, {success:false, message:'학생관리_DB 데이터가 없습니다.'});
    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxStudentId = headers.indexOf('학생ID');
    if (idxStudentId < 0) return outputResult(e, {success:false, message:'학생관리_DB에 학생ID 컬럼이 없습니다.'});

    var idxLearningStartDate = headers.indexOf('학습시작일');
    var idxLearningEndDate = headers.indexOf('학습종료일');
    var idxAttendancePromise = headers.indexOf('출석약속');
    var idxSetPromise = headers.indexOf('세트약속');
    var idxRecipientType = findFirstHeaderIndex_(headers, ['수신자구분','수신구분','수신']);
    var idxLearningAssign = findFirstHeaderIndex_(headers, ['최초배정','학습배정']);
    var idxS4 = headers.indexOf('S4실루엣단계');
    var idxS6 = headers.indexOf('S6테스트모드');
    var idxComplete = headers.indexOf('레벨완료조건');
    var idxNextComplete = headers.indexOf('다음레벨완료조건');
    var idxActiveStatus = headers.indexOf('활성상태');
    var idxSessionStatus = headers.indexOf('세션상태');
    var idxClassName = findFirstHeaderIndex_(headers, ['반명','Class','반']);
    var idxTeacherName = findFirstHeaderIndex_(headers, ['교사명','담당교사']);
    var idxTeacherId = findFirstHeaderIndex_(headers, ['교사ID','담당교사ID']);
    var idxSchool = headers.indexOf('학교');
    var idxGrade = headers.indexOf('학년');
    var idxParentName = headers.indexOf('학부모명');
    var idxParentPhone = headers.indexOf('학부모연락처');

    var targetRow = -1;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][idxStudentId] || '').trim().toUpperCase() === studentId) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) return outputResult(e, {success:false, message:'학생ID를 찾을 수 없습니다: ' + studentId});

    var rowIndex = targetRow - 1;
    var targetRowObject = {};
    headers.forEach(function(header, index){ if (header) targetRowObject[header] = values[rowIndex][index]; });
    if (actor.found && !wmStudentBelongsToActor_(targetRowObject, actor)) {
      return outputResult(e, {success:false, message:'현재 권한으로 수정할 수 없는 학생입니다.'});
    }

    var hasLearningStartDate = hasParam_(['learningStartDate','학습시작일']);
    var hasLearningEndDate = hasParam_(['learningEndDate','학습종료일']);
    var hasAttendancePromise = hasParam_(['attendancePromise','출석약속']);
    var hasSetPromise = hasParam_(['setPromise','세트약속']);
    var hasRecipientType = hasParam_(['recipientType','수신자구분','수신구분','수신']);
    var hasLearningAssign = hasParam_(['learningAssign','최초배정','학습배정','level']);
    var hasS4 = hasParam_(['s4','S4실루엣단계','silhouetteStep']);
    var hasS6 = hasParam_(['s6','S6테스트모드','testMode']);
    var hasComplete = hasParam_(['levelCompleteCondition','레벨완료조건']);
    var hasNextComplete = hasParam_(['nextLevelCompleteCondition','다음레벨완료조건']);
    var hasClassName = hasParam_(['className','반명']);
    var hasTeacherName = hasParam_(['teacherName','교사명']);
    var hasSchool = hasParam_(['school','학교']);
    var hasGrade = hasParam_(['grade','학년']);
    var hasParentName = hasParam_(['parentName','학부모명']);
    var hasParentPhone = hasParam_(['parentPhone','학부모연락처']);
    var hasActiveStatus = hasParam_(['activeStatus','활성상태']);

    function requireColumn_(flag, index, label) {
      if (flag && index < 0) throw new Error('학생관리_DB에 ' + label + ' 컬럼이 없습니다.');
    }
    requireColumn_(hasLearningStartDate, idxLearningStartDate, '학습시작일');
    requireColumn_(hasLearningEndDate, idxLearningEndDate, '학습종료일');
    requireColumn_(hasAttendancePromise, idxAttendancePromise, '출석약속');
    requireColumn_(hasSetPromise, idxSetPromise, '세트약속');
    requireColumn_(hasRecipientType, idxRecipientType, '수신자구분');
    requireColumn_(hasLearningAssign, idxLearningAssign, '최초배정/학습배정');
    requireColumn_(hasS4, idxS4, 'S4실루엣단계');
    requireColumn_(hasS6, idxS6, 'S6테스트모드');
    requireColumn_(hasComplete, idxComplete, '레벨완료조건');
    requireColumn_(hasNextComplete, idxNextComplete, '다음레벨완료조건');
    requireColumn_(hasClassName, idxClassName, '반명');
    requireColumn_(hasTeacherName, idxTeacherName, '교사명');
    requireColumn_(hasSchool, idxSchool, '학교');
    requireColumn_(hasGrade, idxGrade, '학년');
    requireColumn_(hasParentName, idxParentName, '학부모명');
    requireColumn_(hasParentPhone, idxParentPhone, '학부모연락처');
    requireColumn_(hasActiveStatus, idxActiveStatus, '활성상태');

    var previousClassName = idxClassName >= 0 ? String(values[rowIndex][idxClassName] || '').trim() : '';
    var className = hasClassName ? String(p.className || p['반명'] || '').trim() : previousClassName;
    var teacherName = hasTeacherName ? String(p.teacherName || p['교사명'] || '').trim() : (idxTeacherName >= 0 ? String(values[rowIndex][idxTeacherName] || '').trim() : '');
    var school = hasSchool ? String(p.school || p['학교'] || '').trim() : (idxSchool >= 0 ? String(values[rowIndex][idxSchool] || '').trim() : '');
    var grade = hasGrade ? String(p.grade || p['학년'] || '').trim() : (idxGrade >= 0 ? String(values[rowIndex][idxGrade] || '').trim() : '');
    var parentName = hasParentName ? String(p.parentName || p['학부모명'] || '').trim() : (idxParentName >= 0 ? String(values[rowIndex][idxParentName] || '').trim() : '');
    var parentPhone = hasParentPhone ? String(p.parentPhone || p['학부모연락처'] || '').trim() : (idxParentPhone >= 0 ? String(values[rowIndex][idxParentPhone] || '').trim() : '');
    if (hasGrade && grade && !/^(초[1-6]|중[1-3]|고[1-3])$/.test(grade)) {
      return outputResult(e, {success:false, message:'학년은 초1~초6, 중1~중3, 고1~고3 중에서 선택하세요.'});
    }

    var actorAccess = wmStudentDataAccessForActor_(actor);
    if ((hasParentName || hasParentPhone) && (!actor.found || !actorAccess.parent)) {
      return outputResult(e, {success:false, message:'학부모정보 수정권한이 없습니다.'});
    }

    var effectiveStart = idxLearningStartDate >= 0 ? wmNormalizeStudentLearningDate_(values[rowIndex][idxLearningStartDate]) : '';
    var effectiveEnd = idxLearningEndDate >= 0 ? wmNormalizeStudentLearningDate_(values[rowIndex][idxLearningEndDate]) : '';
    if (hasLearningStartDate) {
      effectiveStart = wmNormalizeStudentLearningDate_(p.learningStartDate || p['학습시작일']);
      if (effectiveStart === null) return outputResult(e, {success:false, message:'학습시작일 형식이 올바르지 않습니다.'});
      if (!effectiveStart) return outputResult(e, {success:false, message:'학습시작일을 달력에서 선택해주세요.'});
    }
    if (hasLearningEndDate) {
      effectiveEnd = wmNormalizeStudentLearningDate_(p.learningEndDate || p['학습종료일']);
      if (effectiveEnd === null) return outputResult(e, {success:false, message:'학습종료일 형식이 올바르지 않습니다.'});
    }
    if ((hasLearningStartDate || hasLearningEndDate) && effectiveStart && effectiveEnd && effectiveEnd < effectiveStart) {
      return outputResult(e, {success:false, message:'학습종료일은 학습시작일보다 빠를 수 없습니다.'});
    }

    var attendancePromise = hasAttendancePromise ? String(p.attendancePromise || p['출석약속'] || '').replace(/[^1-7]/g, '').slice(0, 1) : (idxAttendancePromise >= 0 ? String(values[rowIndex][idxAttendancePromise] || '').trim() : '');
    var setPromise = hasSetPromise ? String(p.setPromise || p['세트약속'] || '').replace(/[^1-7]/g, '').slice(0, 1) : (idxSetPromise >= 0 ? String(values[rowIndex][idxSetPromise] || '').trim() : '');
    if (hasAttendancePromise && !attendancePromise) return outputResult(e, {success:false, message:'출석약속은 1~7 중에서 선택하세요.'});
    if (hasSetPromise && !setPromise) return outputResult(e, {success:false, message:'세트약속은 1~7 중에서 선택하세요.'});
    var recipientType = hasRecipientType ? String(p.recipientType || p['수신자구분'] || p['수신구분'] || p['수신'] || '').trim() : (idxRecipientType >= 0 ? String(values[rowIndex][idxRecipientType] || '').trim() : '');

    var assignedTeacherId = hasTeacherName ? String(p.assignedTeacherId || p['교사ID'] || '').trim() : (idxTeacherId >= 0 ? String(values[rowIndex][idxTeacherId] || '').trim() : '');
    if (hasTeacherName) {
      if (teacherName) {
        if (!assignedTeacherId) assignedTeacherId = wmResolveSettingTeacherIdByName_(teacherName);
        if (!assignedTeacherId) return outputResult(e, {success:false, message:'선택한 교사의 교사ID를 찾을 수 없습니다.'});
      } else {
        assignedTeacherId = '';
      }
    }

    var activeStatus = hasActiveStatus ? String(p.activeStatus || p['활성상태'] || '').trim().toUpperCase() : (idxActiveStatus >= 0 ? String(values[rowIndex][idxActiveStatus] || '').trim().toUpperCase() : '');
    if (hasActiveStatus && ['ACTIVE','INACTIVE'].indexOf(activeStatus) === -1) {
      return outputResult(e, {success:false, message:'활성상태는 ACTIVE 또는 INACTIVE만 선택할 수 있습니다.'});
    }
    var sessionStatus = idxSessionStatus >= 0 ? String(values[rowIndex][idxSessionStatus] || 'LOGOUT').trim().toUpperCase() : 'LOGOUT';
    if (['LOGIN','LOGOUT','EXPIRED'].indexOf(sessionStatus) === -1) sessionStatus = 'LOGOUT';

    var learningAssign = hasLearningAssign ? normalizeLearningAssignForLms_(p.learningAssign || p['최초배정'] || p['학습배정'] || p.level || '') : (idxLearningAssign >= 0 ? String(values[rowIndex][idxLearningAssign] || '').trim() : '');
    if (hasLearningAssign && !learningAssign) return outputResult(e, {success:false, message:'최초배정을 선택하세요.'});

    var mode = {
      S4실루엣단계:hasS4 ? normalizeS4SilhouetteStep_(p.s4 || p.S4실루엣단계 || p.silhouetteStep) : (idxS4 >= 0 ? String(values[rowIndex][idxS4] || '').trim() : ''),
      S6테스트모드:hasS6 ? normalizeS6TestMode_(p.s6 || p.S6테스트모드 || p.testMode) : (idxS6 >= 0 ? String(values[rowIndex][idxS6] || '').trim() : ''),
      레벨완료조건:hasComplete ? normalizeLevelCompleteCondition_(p.levelCompleteCondition || p.레벨완료조건) : (idxComplete >= 0 ? String(values[rowIndex][idxComplete] || '').trim() : ''),
      다음레벨완료조건:hasNextComplete ? normalizeLevelCompleteCondition_(p.nextLevelCompleteCondition || p.다음레벨완료조건) : (idxNextComplete >= 0 ? String(values[rowIndex][idxNextComplete] || '').trim() : '')
    };

    var progressRow = null;
    var currentSet = '';
    var nextLevel = '';
    var levelComplete = false;
    if (hasComplete || hasNextComplete) {
      var currentSetMap = WM_CODEGS_ALL_LOCK_V1.STUDENT_ASSIGNMENT.currentSetRead();
      currentSet = String(currentSetMap[studentId] || '').trim();
      progressRow = wmGetCurrentProgressCacheForStudent_(studentId);
      nextLevel = WM_CODEGS_ALL_LOCK_V1.STUDENT_ASSIGNMENT.nextLevelPlusOne(currentSet);
      var progressLevel = Number(String(progressRow && (progressRow['현재레벨'] || progressRow['Set_ID']) || '').match(/\d+/) || 0);
      var currentLevel = Number(String(currentSet || '').match(/\d+/) || 0);
      var levelCompleteText = String(progressRow && progressRow['레벨완료여부'] || '').trim().toUpperCase();
      levelComplete = levelCompleteText === 'Y' || levelCompleteText === 'TRUE' || levelCompleteText === '완료';
      var levelStarted = !!String(progressRow && (progressRow['학습기록ID'] || progressRow['완료Step'] || progressRow['현재Step']) || '').trim();
      if (hasComplete) {
        var savedCurrentCondition = idxComplete >= 0 ? normalizeLevelCompleteCondition_(values[rowIndex][idxComplete]) : '';
        if (mode.레벨완료조건 !== savedCurrentCondition && ((levelStarted && progressLevel && currentLevel && progressLevel === currentLevel) || levelComplete)) {
          return outputResult(e, {success:false, message:levelComplete ? '완료된 레벨의 완료조건은 수정할 수 없습니다.' : '현재 진행 중인 레벨의 완료조건은 수정할 수 없습니다.'});
        }
      }
    }

    if (hasClassName && previousClassName !== className) {
      var classValidation = wmValidateStudentClassMoveForLms_(ss, className);
      if (!classValidation.success) return outputResult(e, {success:false, message:classValidation.message || '반 이동 대상 확인 실패'});
    }

    function writeCellIfChanged_(index, nextValue, numberFormat) {
      if (index < 0) return false;
      var previousValue = String(values[rowIndex][index] || '').trim();
      var nextText = String(nextValue == null ? '' : nextValue).trim();
      if (previousValue === nextText) return false;
      var cell = sheet.getRange(targetRow, index + 1);
      if (nextText) cell.setValue(nextValue); else cell.clearContent();
      if (numberFormat) cell.setNumberFormat(numberFormat);
      values[rowIndex][index] = nextText;
      return true;
    }

    var changedFields = [];
    function mark_(flag, label, changed){ if (flag && changed) changedFields.push(label); }
    mark_(hasLearningStartDate, '학습시작일', writeCellIfChanged_(idxLearningStartDate, effectiveStart ? wmStudentLearningDateValue_(effectiveStart) : '', 'yyyy-MM-dd'));
    mark_(hasLearningEndDate, '학습종료일', writeCellIfChanged_(idxLearningEndDate, effectiveEnd ? wmStudentLearningDateValue_(effectiveEnd) : '', effectiveEnd ? 'yyyy-MM-dd' : ''));
    mark_(hasAttendancePromise, '출석약속', writeCellIfChanged_(idxAttendancePromise, attendancePromise));
    mark_(hasSetPromise, '세트약속', writeCellIfChanged_(idxSetPromise, setPromise));
    mark_(hasRecipientType, '수신자구분', writeCellIfChanged_(idxRecipientType, recipientType));
    mark_(hasLearningAssign, '최초배정', writeCellIfChanged_(idxLearningAssign, learningAssign));
    mark_(hasS4, 'S4실루엣단계', writeCellIfChanged_(idxS4, mode.S4실루엣단계));
    mark_(hasS6, 'S6테스트모드', writeCellIfChanged_(idxS6, mode.S6테스트모드));
    mark_(hasComplete, '레벨완료조건', writeCellIfChanged_(idxComplete, mode.레벨완료조건));
    mark_(hasNextComplete, '다음레벨완료조건', writeCellIfChanged_(idxNextComplete, mode.다음레벨완료조건));
    mark_(hasClassName, '반명', writeCellIfChanged_(idxClassName, className));
    mark_(hasTeacherName, '교사명', writeCellIfChanged_(idxTeacherName, teacherName));
    if (hasTeacherName && idxTeacherId >= 0) {
      var teacherIdCell = sheet.getRange(targetRow, idxTeacherId + 1);
      if (assignedTeacherId) {
        var teacherIdValidation = wmPrepareTeacherIdValidationForLms_(ss, sheet, [teacherIdCell.getA1Notation()]);
        if (!teacherIdValidation.success) return outputResult(e, {success:false, message:teacherIdValidation.message || '교사ID 저장규칙 확인 실패'});
      } else {
        teacherIdCell.clearDataValidations();
      }
      writeCellIfChanged_(idxTeacherId, assignedTeacherId);
    }
    mark_(hasSchool, '학교', writeCellIfChanged_(idxSchool, school));
    mark_(hasGrade, '학년', writeCellIfChanged_(idxGrade, grade));
    mark_(hasParentName, '학부모명', writeCellIfChanged_(idxParentName, parentName));
    mark_(hasParentPhone, '학부모연락처', writeCellIfChanged_(idxParentPhone, parentPhone, '@'));
    mark_(hasActiveStatus, '활성상태', writeCellIfChanged_(idxActiveStatus, activeStatus));
    if (hasNextComplete && levelComplete) writeCellIfChanged_(idxComplete, mode.다음레벨완료조건);

    var classRosterSync = {success:true, skipped:true};
    if (hasClassName && previousClassName !== className) {
      SpreadsheetApp.flush();
      classRosterSync = wmMoveStudentClassRosterForLms_(ss, previousClassName, className, studentId, String(targetRowObject['학생이름'] || studentId));
      if (!classRosterSync.success) {
        return outputResult(e, {success:false, partialSaved:true, message:'학생정보는 저장되었지만 반관리_DB 명단 갱신에 실패했습니다.', error:classRosterSync.message || ''});
      }
    }

    if (hasLearningStartDate || hasLearningEndDate) wmSyncLearningDatesToSheets_();

    var currentProgressStudentInfoSync = {success:true, skipped:true};
    if (hasClassName || hasTeacherName || hasSchool || hasGrade) {
      try {
        currentProgressStudentInfoSync = wmSyncStudentMasterInfoToCurrentProgress_(ss, studentId);
      } catch (syncErr) {
        currentProgressStudentInfoSync = {success:false, updated:false, message:String(syncErr && syncErr.message ? syncErr.message : syncErr)};
      }
    }
    wmClearRuntimeCachesForStudent_(studentId);

    return outputResult(e, {
      success:true,
      message:'학생정보 저장 완료',
      studentId:studentId,
      changedFields:changedFields,
      learningAssign:learningAssign,
      currentSet:currentSet,
      learningStartDate:effectiveStart || '',
      learningEndDate:effectiveEnd || '',
      attendancePromise:attendancePromise,
      setPromise:setPromise,
      recipientType:recipientType,
      className:className,
      teacherName:teacherName,
      school:school,
      grade:grade,
      parentUpdated:hasParentName || hasParentPhone,
      classRosterSync:classRosterSync,
      currentProgressStudentInfoSync:currentProgressStudentInfoSync,
      activeStatus:activeStatus,
      sessionStatus:sessionStatus,
      nextLevel:nextLevel
    });
  } catch (err) {
    return outputResult(e, {success:false, message:'학생정보 저장 오류', error:String(err && err.message ? err.message : err)});
  }
}


/* WM_LMS_CURRENT_PROGRESS_DB_LOOKUP_20260624_V2
 * LMS 학생관리 화면의 현재세트는 1.학생관리_DB가 아니라
 * 8.현재진행_DB의 학생ID별 최신 최종수정일 기준 Set_ID를 그대로 사용합니다.
 */
function buildCurrentProgressSetMapForLms_() {
  var result = {};
  var cacheKey = wmCacheKey_('CP_SET_MAP_EXACT_V1', 'ALL');
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && typeof cached === 'object') return cached;

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('8.현재진행_DB');
    if (!sheet) return result;

    var range = sheet.getDataRange();
    var displayValues = range.getDisplayValues();
    var rawValues = range.getValues();
    if (!displayValues || displayValues.length < 2) return result;

    var headers = displayValues[0].map(function(h) { return String(h || '').trim(); });
    var idxStudentId = headers.indexOf('학생ID');
    var idxSetId = headers.indexOf('Set_ID');
    var idxRecordSetId = headers.indexOf('기록Set_ID');
    var idxUpdated = headers.indexOf('최종수정일');

    if (idxStudentId === -1 || (idxSetId === -1 && idxRecordSetId === -1)) return result;

    function toTime_(rawValue, displayValue, rowIndex) {
      if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        return rawValue.getTime();
      }

      var text = String(displayValue || '').trim();
      var m = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
      if (m) {
        return new Date(
          Number(m[1]),
          Number(m[2]) - 1,
          Number(m[3]),
          Number(m[4] || 0),
          Number(m[5] || 0),
          Number(m[6] || 0)
        ).getTime();
      }

      var parsed = Date.parse(text);
      if (!isNaN(parsed)) return parsed;

      return rowIndex;
    }

    for (var r = 1; r < displayValues.length; r++) {
      var row = displayValues[r] || [];
      var rawRow = rawValues[r] || [];
      var studentId = String(row[idxStudentId] || '').trim().toUpperCase();
      var rawSetId = idxSetId >= 0
        ? String(row[idxSetId] || '').trim()
        : '';
      if (!studentId || !rawSetId) continue;

      var timeValue = idxUpdated >= 0
        ? toTime_(rawRow[idxUpdated], row[idxUpdated], r)
        : r;

      if (!result[studentId] || timeValue >= result[studentId].timeValue) {
        result[studentId] = {
          setId: rawSetId,
          timeValue: timeValue
        };
      }
    }

    Object.keys(result).forEach(function(studentId) {
      result[studentId] = String(result[studentId].setId || '').trim();
    });

    wmCachePutJson_(cacheKey, result, 60);
    return result;
  } catch (err) {
    return result;
  }
}


/* WM_MAP_SINGLE_STUDENT_CURRENT_SET_FAST_20260703_V1
 * 학습맵 진입 시 전체 학생 현재진행맵을 만들지 않고 해당 학생의 최신 현재세트만 읽습니다.
 */
function wmGetCurrentProgressSetForStudent_(studentId) {
  studentId = String(studentId || '').trim().toUpperCase();
  if (!studentId) return '';

  var cached = wmCacheGetJson_(wmCacheKey_('CURRENT_PROGRESS_STUDENT', studentId));
  if (cached) {
    return normalizeCurrentSetForMap_(cached['기록Set_ID'] || '', cached['현재레벨'] || '');
  }

  var row = wmGetCurrentProgressCacheForStudent_(studentId);
  if (!row) return '';
  return normalizeCurrentSetForMap_(row['기록Set_ID'] || '', row['현재레벨'] || '');
}


/* WM_CURRENT_PROGRESS_CACHE_READ_FOR_MAP_20260624_V1
 * Map 표시값은 8.현재진행_DB의 본문 컬럼을 1순위로 사용합니다.
 * - 세트진행률: 현재 학습 세트 Step 진행률
 * - 레벨진행률: 완료세트수 / 전체세트수
 * 최근30건JSON은 상세 기록용이며 진행률 표시 기준으로 사용하지 않습니다.
 */
function wmGetCurrentProgressCacheForStudent_(studentId) {
  /* WM_CURRENT_PROGRESS_DIRECT_ONE_ROW_V1_20260821
   * 8.현재진행_DB는 학생ID당 현재진행 1행이 공식 기준입니다.
   * 동일학생 전체검색/최종수정일 비교를 하지 않고 학생ID 1건을 바로 찾아 그 1행만 읽습니다.
   */
  var result = null;
  studentId = String(studentId || '').trim().toUpperCase();
  if (!studentId) return result;

  var cacheKey = wmCacheKey_('CURRENT_PROGRESS_STUDENT', studentId);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && cached['학생ID']) return cached;

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss && ss.getSheetByName('8.현재진행_DB');
    if (!sheet) return result;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return result;

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });
    var idxStudent = headers.indexOf('학생ID');
    if (idxStudent < 0) return result;

    var rowNumber = wmGetStudentRowNumberByIdIndex_(
      sheet,
      idxStudent + 1,
      lastRow,
      'CURRENT_PROGRESS_ROW_INDEX',
      studentId
    );
    if (!rowNumber) return result;
    var rowDisplay = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
    if (wmStudentIdCompareKey_(rowDisplay[idxStudent]) !== wmStudentIdCompareKey_(studentId)) {
      rowNumber = wmGetStudentRowNumberByIdIndex_(
        sheet,
        idxStudent + 1,
        lastRow,
        'CURRENT_PROGRESS_ROW_INDEX',
        studentId,
        true
      );
      if (!rowNumber) return result;
      rowDisplay = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
      if (wmStudentIdCompareKey_(rowDisplay[idxStudent]) !== wmStudentIdCompareKey_(studentId)) return result;
    }
    var obj = buildRecordObjectByHeaders_(headers, rowDisplay);
    obj.__rowNumber = rowNumber;

    wmCachePutJson_(cacheKey, obj, 60);
    return obj;
  } catch (err) {
    return null;
  }
}


/* WM_STUDENT_MASTER_TO_CURRENT_PROGRESS_SYNC_V1_20260821
 * 1.학생관리_DB의 학생 기본정보가 바뀌면 동일 학생ID의 8.현재진행_DB 기본정보도 즉시 동기화합니다.
 * 과거 2.학습기록_DB는 수정하지 않으며, 진행값/점수/최종수정일도 건드리지 않습니다. */
function wmSyncStudentMasterInfoToCurrentProgress_(ss, studentId) {
  studentId = String(studentId || '').trim();
  if (!studentId) return {success:false, updated:false, message:'학생ID 없음'};
  var batch = wmSyncStudentMasterInfoToCurrentProgressBatch_(ss, [studentId]);
  if (!batch || batch.success !== true) {
    return {
      success:false,
      updated:false,
      message:String(batch && batch.message || '현재진행_DB 동기화 실패')
    };
  }
  var one = batch.results && batch.results.length ? batch.results[0] : null;
  return {
    success:true,
    updated:!!(one && one.updated),
    updatedCells:Number(one && one.updatedCells || 0) || 0,
    message:String(one && one.message || '')
  };
}

/* WM_STUDENT_MASTER_TO_CURRENT_PROGRESS_BATCH_SYNC_V1_20260821
 * 반배정처럼 여러 학생을 한 번에 수정하는 서버 저장경로에서도
 * 각 학생ID의 현재진행_DB 기본정보를 같은 요청 안에서 즉시 동기화합니다. */
function wmSyncStudentMasterInfoToCurrentProgressBatch_(ss, studentIds) {
  var ids = [];
  var seen = {};
  (studentIds || []).forEach(function(id) {
    var text = String(id || '').trim();
    var key = text.toUpperCase();
    if (!text || seen[key]) return;
    seen[key] = true;
    ids.push(text);
  });
  if (!ids.length) return {success:true, requested:0, updatedStudents:0, updatedCells:0, results:[]};

  ss = ss || getLmsSpreadsheet_();
  var studentSheet = ss && ss.getSheetByName('1.학생관리_DB');
  var progressSheet = ss && ss.getSheetByName('8.현재진행_DB');
  if (!studentSheet || !progressSheet) {
    return {success:false, requested:ids.length, updatedStudents:0, updatedCells:0, results:[], message:'학생관리_DB 또는 현재진행_DB 없음'};
  }

  var studentLastRow = studentSheet.getLastRow();
  var studentLastCol = studentSheet.getLastColumn();
  var progressLastRow = progressSheet.getLastRow();
  var progressLastCol = progressSheet.getLastColumn();
  if (studentLastRow < 2 || studentLastCol < 1 || progressLastRow < 2 || progressLastCol < 1) {
    return {success:true, requested:ids.length, updatedStudents:0, updatedCells:0, results:ids.map(function(id){return {studentId:id, updated:false, updatedCells:0, message:'동기화 대상 행 없음'};})};
  }

  var wanted = {};
  ids.forEach(function(id){ wanted[String(id).trim().toUpperCase()] = id; });

  var studentValues = studentSheet.getRange(1, 1, studentLastRow, studentLastCol).getDisplayValues();
  var studentHeaders = studentValues[0].map(function(h){ return String(h || '').trim(); });
  var sIdxId = studentHeaders.indexOf('학생ID');
  if (sIdxId < 0) return {success:false, requested:ids.length, updatedStudents:0, updatedCells:0, results:[], message:'학생관리_DB 학생ID 컬럼 없음'};

  function studentIndex_(names) {
    for (var i = 0; i < names.length; i++) {
      var idx = studentHeaders.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }
  var sourceIndexes = {
    '학생이름':studentIndex_(['학생이름','학생명','이름']),
    '학교':studentIndex_(['학교','학교명']),
    '학년':studentIndex_(['학년']),
    'Class':studentIndex_(['반명','Class','반']),
    '교사명':studentIndex_(['교사명','담당교사'])
  };
  var sourceById = {};
  for (var sr = 1; sr < studentValues.length; sr++) {
    var sid = String(studentValues[sr][sIdxId] || '').trim();
    var key = sid.toUpperCase();
    if (!key || !wanted[key] || sourceById[key]) continue;
    var source = {};
    Object.keys(sourceIndexes).forEach(function(header){
      var idx = sourceIndexes[header];
      source[header] = idx >= 0 ? String(studentValues[sr][idx] || '').trim() : '';
    });
    sourceById[key] = source;
  }

  var progressValues = progressSheet.getRange(1, 1, progressLastRow, progressLastCol).getDisplayValues();
  var progressHeaders = progressValues[0].map(function(h){ return String(h || '').trim(); });
  var pIdxId = progressHeaders.indexOf('학생ID');
  if (pIdxId < 0) return {success:false, requested:ids.length, updatedStudents:0, updatedCells:0, results:[], message:'현재진행_DB 학생ID 컬럼 없음'};

  function columnA1_(columnNumber) {
    var n = Number(columnNumber || 0);
    var out = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  var changesByGroup = {};
  var perStudentCount = {};
  var foundProgress = {};
  for (var pr = 1; pr < progressValues.length; pr++) {
    var pid = String(progressValues[pr][pIdxId] || '').trim();
    var pKey = pid.toUpperCase();
    if (!pKey || !wanted[pKey]) continue;
    foundProgress[pKey] = true;
    var source = sourceById[pKey];
    if (!source) continue;
    Object.keys(source).forEach(function(header){
      var colIndex = progressHeaders.indexOf(header);
      if (colIndex < 0) return;
      var oldValue = String(progressValues[pr][colIndex] || '').trim();
      var newValue = String(source[header] || '').trim();
      if (oldValue === newValue) return;
      var groupKey = String(colIndex) + '\u0001' + newValue;
      if (!changesByGroup[groupKey]) changesByGroup[groupKey] = {colIndex:colIndex, value:newValue, a1:[]};
      changesByGroup[groupKey].a1.push(columnA1_(colIndex + 1) + String(pr + 1));
      perStudentCount[pKey] = (perStudentCount[pKey] || 0) + 1;
      progressValues[pr][colIndex] = newValue;
    });
  }

  Object.keys(changesByGroup).forEach(function(key){
    var group = changesByGroup[key];
    if (!group.a1.length) return;
    var ranges = progressSheet.getRangeList(group.a1);
    if (group.value) ranges.setValue(group.value);
    else ranges.clearContent();
  });

  var updatedStudents = 0;
  var updatedCells = 0;
  var results = ids.map(function(id){
    var key = String(id || '').trim().toUpperCase();
    var count = Number(perStudentCount[key] || 0) || 0;
    if (count > 0) updatedStudents++;
    updatedCells += count;
    var message = !sourceById[key] ? '학생관리_DB 학생 없음' : (!foundProgress[key] ? '현재진행_DB 행 없음' : '');
    return {studentId:id, updated:count > 0, updatedCells:count, message:message};
  });

  wmClearRuntimeCachesForStudents_(ids);
  return {success:true, requested:ids.length, updatedStudents:updatedStudents, updatedCells:updatedCells, results:results};
}

function wmParseCurrentProgressPercent_(value) {
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return 0;
  var n = Number(text.replace(/%/g, '').replace(/,/g, ''));
  if (!isFinite(n)) return 0;
  if (text.indexOf('%') >= 0) return Math.max(0, Math.min(100, n));
  return Math.max(0, Math.min(100, n));
}

function wmFormatCurrentProgressPercentText_(value) {
  var n = wmParseCurrentProgressPercent_(value);
  return String(Math.max(0, Math.min(100, Math.round(n)))) + '%';
}

function wmNormalizeCurrentProgressRoundValue_(value) {
  /* WM_CURRENT_PROGRESS_ROUND_VALUE_GUARD_20260707_V1
   * 순차완주회차/히스토리횟수는 진행률(%)이 아니라 회차 숫자만 저장합니다.
   * - 학습없음: 0
   * - 1회차 진행/완료: 1
   * - 2회차 진행/완료: 2
   * - 3회차 진행/완료 이상: 3
   */
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return 0;
  var match = text.match(/\d+/);
  var n = match ? Number(match[0]) : Number(value);
  if (!isFinite(n) || n <= 0) return 0;
  return Math.max(1, Math.min(3, Math.floor(n)));
}


function wmNormalizeHeaderKeyForCurrentProgress_(header) {
  return String(header === undefined || header === null ? '' : header)
    .replace(/\s+/g, '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();
}

function wmForceCurrentProgressRowFormulaValues_(progressHeaders, row, cacheFields, record, existingRecentScore) {
  /* WM_CURRENT_PROGRESS_RIGHT_COLUMN_FINAL_GUARD_20260707_V2
   * 현재진행_DB 저장 직전 최종 방어입니다.
   * - 순차완주회차에는 진행률(%)이 아니라 0/1/2/3 회차 숫자만 씁니다.
   * - 순차완주세트에는 현재 진행중 세트를 제외한 순차 완료 세트 수만 씁니다.
   * - 최근점수에는 레벨값(예: 5레벨)이 들어가지 못하게 막습니다.
   */
  progressHeaders = progressHeaders || [];
  row = row || [];
  cacheFields = cacheFields || {};
  record = record || {};

  function setByHeader(headerName, value) {
    var target = wmNormalizeHeaderKeyForCurrentProgress_(headerName);
    for (var i = 0; i < progressHeaders.length; i++) {
      if (wmNormalizeHeaderKeyForCurrentProgress_(progressHeaders[i]) === target) {
        row[i] = value;
      }
    }
  }

  setByHeader('레벨완료횟수', String(wmNormalizeCurrentProgressRoundValue_(cacheFields['레벨완료횟수'])));
  setByHeader('순차완주세트', wmNormalizeCurrentProgressSetCountValue_(cacheFields['순차완주세트']));
  setByHeader('순차완주회차', String(wmNormalizeCurrentProgressRoundValue_(cacheFields['순차완주회차'])));
  setByHeader('완료세트수', Number(cacheFields['완료세트수'] || 0));
  setByHeader('전체세트수', Number(cacheFields['전체세트수'] || 0));
  setByHeader('세트진행률', wmFormatCurrentProgressPercentText_(cacheFields['세트진행률']));
  setByHeader('레벨진행률', wmFormatCurrentProgressPercentText_(cacheFields['레벨진행률']));

  var safeRecentScore = wmExtractScoreForCurrentProgress_({
    '점수': cacheFields['최근점수'] || record['점수'],
    '한영주관식': record['한영주관식'],
    '최근점수': existingRecentScore
  });
  setByHeader('최근점수', safeRecentScore || '');

  return row;
}

function wmNormalizeCurrentProgressHistoryRoundsValue_(value) {
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return '';
  return text.split('|').map(function(part) {
    return wmNormalizeCurrentProgressRoundValue_(part);
  }).join('|');
}

function wmNormalizeCurrentProgressSetCountValue_(value) {
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return 0;
  var match = text.match(/\d+/);
  var n = match ? Number(match[0]) : Number(value);
  if (!isFinite(n) || n < 0) return 0;
  return Math.max(0, Math.floor(n));
}

function wmApplyCurrentProgressCacheToOfficialLevelProgress_(officialLevelProgress, currentProgressCache) {
  if (!officialLevelProgress || !currentProgressCache) return officialLevelProgress;

  var completedSetCount = Number(currentProgressCache['완료세트수'] || 0);
  var totalSetCount = Number(currentProgressCache['전체세트수'] || 0);
  var levelProgressPercent = wmParseCurrentProgressPercent_(currentProgressCache['레벨진행률'] || currentProgressCache['진행률'] || 0);
  var setProgressPercent = wmParseCurrentProgressPercent_(currentProgressCache['세트진행률'] || 0);

  if (totalSetCount > 0) {
    if (!levelProgressPercent && completedSetCount > 0) {
      levelProgressPercent = Math.round((completedSetCount / totalSetCount) * 100);
    }

    officialLevelProgress.완료세트수 = completedSetCount;
    officialLevelProgress.전체세트수 = totalSetCount;
    officialLevelProgress.레벨진행률 = levelProgressPercent;
    officialLevelProgress.세트진행률 = setProgressPercent;

    /* 기존 Map.html 호환 별칭: 진행률/progressPercent는 레벨진행률로 유지합니다. */
    officialLevelProgress.진행률 = levelProgressPercent;
    officialLevelProgress.progressPercent = levelProgressPercent;
    officialLevelProgress.levelProgressPercent = levelProgressPercent;
    officialLevelProgress.stepProgressPercent = setProgressPercent;
    officialLevelProgress.currentProgressCacheSource = '8.현재진행_DB';
  }

  return officialLevelProgress;
}


/* WM_REAL_MAP_CURRENT_PROGRESS_FALLBACK_20260625_V1
 * 리얼부스에서 1.학생관리_DB 일부 값은 읽히지만 2.학습기록_DB/8.현재진행_DB 대표값이 화면에 반영되지 않는 경우를 막기 위해,
 * 8.현재진행_DB의 최신 1행을 학습맵 표시용 studentProfile / setStatusMap / progress 객체에 직접 병합합니다.
 * TEST/REAL 전환 구조는 건드리지 않고, 기존 WM_ENV / WM_ENV_CONFIG를 그대로 사용합니다.
 */
function wmNormalizeCurrentProgressSetIdForMap_(currentProgressCache, fallbackSetId) {
  /* WM_MAP_CURRENT_PROGRESS_RECORD_SET_ONLY_20260820_V1
   * 현재 학습 세트는 8.현재진행_DB 기록Set_ID 한 컬럼만 사용합니다. */
  currentProgressCache = currentProgressCache || {};
  var raw = String(currentProgressCache['기록Set_ID'] || '').trim();
  if (!raw) return '';
  return normalizeCurrentSetForMap_(raw, currentProgressCache['현재레벨'] || '');
}

function wmApplyCurrentProgressCacheToStudentProfile_(studentProfile, currentProgressCache) {
  studentProfile = studentProfile || {};
  currentProgressCache = currentProgressCache || null;
  if (!currentProgressCache) return studentProfile;

  function keepValue_(key, value) {
    var text = String(value || '').trim();
    if (text) studentProfile[key] = text;
  }

  if (!String(studentProfile['학생ID'] || '').trim()) keepValue_('학생ID', currentProgressCache['학생ID']);
  if (!String(studentProfile['학생이름'] || '').trim()) keepValue_('학생이름', currentProgressCache['학생이름']);
  if (!String(studentProfile['학교'] || '').trim()) keepValue_('학교', currentProgressCache['학교']);
  if (!String(studentProfile['학년'] || '').trim()) keepValue_('학년', currentProgressCache['학년']);
  if (!String(studentProfile['Class'] || '').trim()) keepValue_('Class', currentProgressCache['Class']);
  if (!String(studentProfile['교사명'] || '').trim()) keepValue_('교사명', currentProgressCache['교사명']);

  var levelText = String(currentProgressCache['현재레벨'] || '').trim();
  if (levelText) {
    if (/^[0-9]+$/.test(levelText)) levelText = levelText + '레벨';
    studentProfile['현재레벨'] = levelText;
    studentProfile.currentLevel = levelText;
  }

  var cacheSetId = wmNormalizeCurrentProgressSetIdForMap_(currentProgressCache, studentProfile['현재세트']);
  if (cacheSetId) {
    studentProfile['현재세트'] = cacheSetId;
    studentProfile.currentSet = cacheSetId;
  }

  var learningMode = studentProfile['학습모드'] || buildDefaultLearningMode_();
  if (currentProgressCache['레벨완료조건']) learningMode.레벨완료조건 = normalizeLevelCompleteCondition_(currentProgressCache['레벨완료조건']);
  if (currentProgressCache['레벨회차추가']) learningMode.레벨회차추가 = normalizeLevelExtraRounds_(currentProgressCache['레벨회차추가']);
  if (currentProgressCache['S4실루엣단계']) learningMode.S4실루엣단계 = normalizeS4SilhouetteStep_(currentProgressCache['S4실루엣단계']);
  if (currentProgressCache['S6테스트모드']) learningMode.S6테스트모드 = normalizeS6TestMode_(currentProgressCache['S6테스트모드']);
  studentProfile['학습모드'] = learningMode;
  studentProfile['레벨완료조건'] = learningMode.레벨완료조건;
  studentProfile['레벨완료횟수'] = wmNormalizeCurrentProgressRoundValue_(currentProgressCache['레벨완료횟수']);
  studentProfile['순차완주세트'] = String(currentProgressCache['순차완주세트'] || '').trim();
  studentProfile['순차완주회차'] = wmNormalizeCurrentProgressRoundValue_(currentProgressCache['순차완주회차']);
  studentProfile['히스토리레벨'] = String(currentProgressCache['히스토리레벨'] || currentProgressCache['완료된레벨'] || '').trim();
  studentProfile['히스토리횟수'] = String(currentProgressCache['히스토리횟수'] || currentProgressCache['완료횟수'] || '').trim();
  studentProfile['레벨회차추가'] = learningMode.레벨회차추가;
  studentProfile['목표회차'] = buildLevelTargetRoundGoalFromLearningMode_(learningMode).목표회차;

  return studentProfile;
}

function wmBuildCurrentProgressLearningStateForMap_(studentProfile, currentProgressCache) {
  studentProfile = studentProfile || {};
  currentProgressCache = currentProgressCache || {};

  var currentSet = wmNormalizeCurrentProgressSetIdForMap_(currentProgressCache, '');
  var level = extractLevelNumberForMap_(currentProgressCache['현재레벨'] || '');
  var completedSetCount = Number(currentProgressCache['완료세트수'] || 0);
  var totalSetCount = Number(currentProgressCache['전체세트수'] || 0);
  if (!totalSetCount) totalSetCount = getLevelSetSequence_(level).length;

  var levelProgressPercent = wmParseCurrentProgressPercent_(currentProgressCache['레벨진행률'] || 0);
  if (!levelProgressPercent && totalSetCount > 0 && completedSetCount > 0) {
    levelProgressPercent = Math.round((completedSetCount / totalSetCount) * 100);
  }

  var setProgressPercent = wmParseCurrentProgressPercent_(currentProgressCache['세트진행률'] || 0);
  var completedRoundText = String(currentProgressCache['순차완주회차'] || currentProgressCache['완료회차'] || '').trim();
  var officialRound = wmNormalizeCurrentProgressRoundValue_(currentProgressCache['순차완주회차'] || currentProgressCache['완료회차'] || 0);
  var currentStep = String(currentProgressCache['현재Step'] || '').trim();
  var completedStep = String(currentProgressCache['완료Step'] || '').trim();

  var setStatusMap = {};
  if (currentSet) {
    setStatusMap[currentSet] = {
      setId: currentSet,
      repeatCount: officialRound,
      officialRound: officialRound,
      latestStatus: setProgressPercent >= 100 ? '완료' : '학습중',
      latestScore: String(currentProgressCache['최근점수'] || '').trim(),
      latestDate: String(currentProgressCache['최종수정일'] || '').trim(),
      currentProgressStep: currentStep,
      lastCompletedStep: completedStep,
      completedStepCount: Math.max(0, Math.min(6, Math.floor((setProgressPercent || 0) / 16.6))),
      progressPercent: setProgressPercent,
      progressRank: 0
    };
    setStatusMap['WM' + currentSet] = setStatusMap[currentSet];
  }

  var officialLevelProgress = {
    level: level,
    현재레벨: level,
    완료세트수: completedSetCount,
    전체세트수: totalSetCount,
    레벨진행률: levelProgressPercent,
    세트진행률: setProgressPercent,
    진행률: levelProgressPercent,
    progressPercent: levelProgressPercent,
    levelProgressPercent: levelProgressPercent,
    stepProgressPercent: setProgressPercent,
    레벨완료: String(currentProgressCache['레벨완료여부'] || '').trim() === '완료' || String(currentProgressCache['레벨완료여부'] || '').toUpperCase() === 'TRUE'
  };

  return {
    currentSet: currentSet,
    setStatusMap: setStatusMap,
    officialLevelProgress: officialLevelProgress,
    currentLearningProgress: {
      현재세트: currentSet,
      currentSet: currentSet,
      현재Step: currentStep,
      완료Step: completedStep,
      세트진행률: setProgressPercent,
      stepProgressPercent: setProgressPercent,
      최근점수: String(currentProgressCache['최근점수'] || '').trim()
    }
  };
}


/* WM_LMS_STUDENT_REGISTER_API_20260721_V1
 * 학생등록 전용 빠른 중복확인 및 1.학생관리_DB 신규행 저장.
 * 기존 학생조회/학습/LOCK 로직은 변경하지 않습니다.
 */
function wmCheckStudentIdForLms_(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var studentId = String(p.studentId || p['학생ID'] || '').trim();
    if (!/^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]{4,10}$/.test(studentId)) {
      return {success:false, available:false, message:'학생ID는 영문+숫자 4~10자입니다.'};
    }
    var sheet = getLmsSpreadsheet_().getSheetByName('1.학생관리_DB');
    if (!sheet) return {success:false, available:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {success:true, available:true, message:'사용 가능한 학생ID입니다.'};
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
    var idCol = headers.indexOf('학생ID');
    if (idCol < 0) return {success:false, available:false, message:'학생ID 컬럼을 찾을 수 없습니다.'};
    var wanted = studentId.toUpperCase();
    var ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toUpperCase() === wanted) {
        return {success:true, available:false, message:'이미 사용 중인 학생ID입니다.'};
      }
    }
    return {success:true, available:true, message:'사용 가능한 학생ID입니다.'};
  } catch (err) {
    return {success:false, available:false, message:'학생ID 중복확인 오류', error:String(err && err.message ? err.message : err)};
  }
}


/* WM_STUDENT_REGISTER_CLASS_SYNC_V1
 * 학생관리에서 학생을 신규등록할 때 선택한 반의 7-2.반관리_DB 명단/학생수도 즉시 갱신합니다.
 * 반ID/교사일련번호는 사용하지 않습니다.
 */
/* WM_ADMIN_SAVE_INTEGRITY_HELPERS_V1_20260821
 * 학생/반/교사 관리 저장은 본 저장과 관계없는 항목 때문에 실패하지 않도록
 * 저장 전 대상 검증, 반 명단 재구성, 반 담당교사 일괄 동기화를 분리합니다. */
function wmValidateStudentClassMoveForLms_(ss, newClassName) {
  newClassName = String(newClassName || '').trim();
  if (!newClassName) return {success:true, skipped:true};
  ss = ss || getLmsSpreadsheet_();
  var sheet = ss && ss.getSheetByName('7-2.반관리_DB');
  if (!sheet) return {success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {success:false, message:'7-2.반관리_DB에 등록된 반이 없습니다.'};
  var values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function(v){ return String(v || '').trim(); });
  var idxClass = headers.indexOf('반명');
  var idxIds = headers.indexOf('학생ID목록');
  var idxNames = headers.indexOf('학생이름목록');
  var idxCount = headers.indexOf('학생수');
  if (idxClass < 0 || idxIds < 0 || idxNames < 0 || idxCount < 0) {
    return {success:false, message:'7-2.반관리_DB 필수 컬럼을 확인하세요.'};
  }
  var key = newClassName.toUpperCase();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idxClass] || '').trim().toUpperCase() === key) return {success:true, row:r + 1};
  }
  return {success:false, message:'선택한 반을 7-2.반관리_DB에서 찾을 수 없습니다: ' + newClassName};
}

function wmRebuildClassRostersFromStudentDb_(ss, classNames) {
  ss = ss || getLmsSpreadsheet_();
  var classSheet = ss && ss.getSheetByName('7-2.반관리_DB');
  var studentSheet = ss && ss.getSheetByName('1.학생관리_DB');
  if (!classSheet || !studentSheet) return {success:false, message:'학생관리_DB 또는 반관리_DB 없음'};

  var targets = {};
  (classNames || []).forEach(function(name){
    var text = String(name || '').trim();
    if (text) targets[text.toUpperCase()] = text;
  });
  var targetKeys = Object.keys(targets);
  if (!targetKeys.length) return {success:true, updatedClasses:0};

  var classValues = classSheet.getDataRange().getDisplayValues();
  var classHeaders = classValues[0].map(function(v){ return String(v || '').trim(); });
  var cIdxClass = classHeaders.indexOf('반명');
  var cIdxIds = classHeaders.indexOf('학생ID목록');
  var cIdxNames = classHeaders.indexOf('학생이름목록');
  var cIdxCount = classHeaders.indexOf('학생수');
  var cIdxUpdated = classHeaders.indexOf('수정일');
  if (cIdxClass < 0 || cIdxIds < 0 || cIdxNames < 0 || cIdxCount < 0) return {success:false, message:'반관리_DB 명단 컬럼 누락'};

  var studentValues = studentSheet.getDataRange().getDisplayValues();
  var studentHeaders = studentValues[0].map(function(v){ return String(v || '').trim(); });
  var sIdxId = studentHeaders.indexOf('학생ID');
  var sIdxName = findFirstHeaderIndex_(studentHeaders, ['학생이름','학생명','이름']);
  var sIdxClass = findFirstHeaderIndex_(studentHeaders, ['반명','Class','반','현재반','클래스']);
  if (sIdxId < 0 || sIdxClass < 0) return {success:false, message:'학생관리_DB 학생ID 또는 반명 컬럼 누락'};

  var rosterByKey = {};
  targetKeys.forEach(function(key){ rosterByKey[key] = {ids:[], names:[], seen:{}}; });
  for (var sr = 1; sr < studentValues.length; sr++) {
    var classKey = String(studentValues[sr][sIdxClass] || '').trim().toUpperCase();
    if (!rosterByKey[classKey]) continue;
    var id = String(studentValues[sr][sIdxId] || '').trim();
    if (!id) continue;
    var idKey = id.toUpperCase();
    if (rosterByKey[classKey].seen[idKey]) continue;
    rosterByKey[classKey].seen[idKey] = true;
    rosterByKey[classKey].ids.push(id);
    rosterByKey[classKey].names.push(String(sIdxName >= 0 ? studentValues[sr][sIdxName] : '').trim() || id);
  }

  var updatedClasses = 0;
  var now = new Date();
  for (var cr = 1; cr < classValues.length; cr++) {
    var key = String(classValues[cr][cIdxClass] || '').trim().toUpperCase();
    if (!rosterByKey[key]) continue;
    var roster = rosterByKey[key];
    classSheet.getRange(cr + 1, cIdxIds + 1).setValue(roster.ids.join(', '));
    classSheet.getRange(cr + 1, cIdxNames + 1).setValue(roster.names.join(', '));
    classSheet.getRange(cr + 1, cIdxCount + 1).setValue(roster.ids.length);
    if (cIdxUpdated >= 0) classSheet.getRange(cr + 1, cIdxUpdated + 1).setValue(now);
    updatedClasses++;
  }
  return {success:true, updatedClasses:updatedClasses};
}

function wmApplyClassTeacherChangeForLms_(ss, className, teacherId, teacherName, options) {
  options = options || {};
  ss = ss || getLmsSpreadsheet_();
  className = String(className || '').trim();
  teacherId = String(teacherId || '').trim();
  teacherName = String(teacherName || '').trim();
  var unassigned = !teacherName || teacherName === '추후선택';
  var studentTeacherName = unassigned ? '' : teacherName;
  var studentTeacherId = unassigned ? '' : teacherId;
  if (!className) return {success:false, message:'반명이 없습니다.'};

  var classSheet = ss.getSheetByName('7-2.반관리_DB');
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  if (!classSheet || !studentSheet) return {success:false, message:'학생관리_DB 또는 반관리_DB 없음'};

  var classValues = classSheet.getDataRange().getDisplayValues();
  var classHeaders = classValues[0].map(function(v){ return String(v || '').trim(); });
  var cIdxClass = classHeaders.indexOf('반명');
  var cIdxTeacherId = classHeaders.indexOf('담당교사ID');
  var cIdxTeacherName = classHeaders.indexOf('담당교사명');
  var cIdxUpdated = classHeaders.indexOf('수정일');
  if (cIdxClass < 0 || cIdxTeacherId < 0 || cIdxTeacherName < 0) return {success:false, message:'반관리_DB 담당교사 컬럼 누락'};

  var classRow = 0;
  var key = className.toUpperCase();
  for (var cr = 1; cr < classValues.length; cr++) {
    if (String(classValues[cr][cIdxClass] || '').trim().toUpperCase() === key) { classRow = cr + 1; break; }
  }
  if (!classRow) return {success:false, message:'반관리_DB에서 반을 찾을 수 없습니다: ' + className};

  var previousTeacherId = Object.prototype.hasOwnProperty.call(options, 'previousTeacherId')
    ? String(options.previousTeacherId || '').trim()
    : String(classValues[classRow - 1][cIdxTeacherId] || '').trim();
  var previousTeacherName = Object.prototype.hasOwnProperty.call(options, 'previousTeacherName')
    ? String(options.previousTeacherName || '').trim()
    : String(classValues[classRow - 1][cIdxTeacherName] || '').trim();

  if (options.updateClassRow !== false) {
    classSheet.getRange(classRow, cIdxTeacherId + 1).setValue(studentTeacherId);
    classSheet.getRange(classRow, cIdxTeacherName + 1).setValue(unassigned ? '추후선택' : teacherName);
    if (cIdxUpdated >= 0) classSheet.getRange(classRow, cIdxUpdated + 1).setValue(new Date());
  }

  var studentValues = studentSheet.getDataRange().getDisplayValues();
  var studentHeaders = studentValues[0].map(function(v){ return String(v || '').trim(); });
  var sIdxId = studentHeaders.indexOf('학생ID');
  var sIdxClass = findFirstHeaderIndex_(studentHeaders, ['반명','Class','반','현재반','클래스']);
  var sIdxTeacherName = findFirstHeaderIndex_(studentHeaders, ['교사명','담당교사','현재담당교사']);
  var sIdxTeacherId = findFirstHeaderIndex_(studentHeaders, ['교사ID','담당교사ID','현재담당교사ID']);
  if (sIdxId < 0 || sIdxClass < 0 || sIdxTeacherName < 0) return {success:false, message:'학생관리_DB 반/교사 컬럼 누락'};

  var studentIds = [];
  var teacherNameA1 = [];
  var teacherIdA1 = [];
  function colA1_(n){ var out=''; while(n>0){ var rem=(n-1)%26; out=String.fromCharCode(65+rem)+out; n=Math.floor((n-1)/26); } return out; }
  for (var sr = 1; sr < studentValues.length; sr++) {
    if (String(studentValues[sr][sIdxClass] || '').trim().toUpperCase() !== key) continue;
    var sid = String(studentValues[sr][sIdxId] || '').trim();
    if (sid) studentIds.push(sid);
    teacherNameA1.push(colA1_(sIdxTeacherName + 1) + String(sr + 1));
    if (sIdxTeacherId >= 0) teacherIdA1.push(colA1_(sIdxTeacherId + 1) + String(sr + 1));
  }
  if (teacherNameA1.length) {
    var nameRanges = studentSheet.getRangeList(teacherNameA1);
    if (studentTeacherName) nameRanges.setValue(studentTeacherName); else nameRanges.clearContent();
  }
  if (teacherIdA1.length) {
    var idRanges = studentSheet.getRangeList(teacherIdA1);
    if (studentTeacherId) {
      var teacherIdValidation = wmPrepareTeacherIdValidationForLms_(ss, studentSheet, teacherIdA1);
      if (!teacherIdValidation.success) return {success:false, message:teacherIdValidation.message || '학생 교사ID 저장규칙 확인 실패'};
      idRanges.setValue(studentTeacherId);
    } else {
      idRanges.clearDataValidations();
      idRanges.clearContent();
    }
  }

  var teacherSheet = wmGetSettingCenterSheet_();
  var teacherHeaders = wmEnsureSettingCenterHeader_(teacherSheet);
  var teacherValues = teacherSheet.getDataRange().getDisplayValues();
  var tIdxId = wmFindHeaderIndex_(teacherHeaders, ['교사ID']);
  var tIdxName = wmFindHeaderIndex_(teacherHeaders, ['교사명','이름']);
  var tIdxClasses = wmFindHeaderIndex_(teacherHeaders, ['담당반목록','담당반']);
  var tIdxTotal = wmFindHeaderIndex_(teacherHeaders, ['총담당반']);
  var tIdxUpdated = wmFindHeaderIndex_(teacherHeaders, ['수정일','수정일시']);
  if (tIdxId >= 0 && tIdxName >= 0 && tIdxClasses >= 0) {
    var oldIdKey = previousTeacherId.toUpperCase();
    var oldNameKey = previousTeacherName.toUpperCase();
    var newIdKey = studentTeacherId.toUpperCase();
    var newNameKey = studentTeacherName.toUpperCase();
    for (var tr = 1; tr < teacherValues.length; tr++) {
      var rowId = String(teacherValues[tr][tIdxId] || '').trim();
      var rowName = String(teacherValues[tr][tIdxName] || '').trim();
      var isOld = (!!oldIdKey && rowId.toUpperCase() === oldIdKey) || (!!oldNameKey && rowName.toUpperCase() === oldNameKey);
      var isNew = (!!newIdKey && rowId.toUpperCase() === newIdKey) || (!!newNameKey && rowName.toUpperCase() === newNameKey);
      if (!isOld && !isNew) continue;
      var classes = wmNormalizeSettingClassList_(teacherValues[tr][tIdxClasses]);
      classes = classes.filter(function(name){ return String(name || '').trim().toUpperCase() !== key; });
      if (isNew) classes.push(className);
      classes = wmNormalizeSettingClassList_(classes.join(', '));
      teacherSheet.getRange(tr + 1, tIdxClasses + 1).setValue(classes.length ? classes.join(', ') : '-');
      if (tIdxTotal >= 0) teacherSheet.getRange(tr + 1, tIdxTotal + 1).setValue(classes.length ? classes.length : '-');
      if (tIdxUpdated >= 0) teacherSheet.getRange(tr + 1, tIdxUpdated + 1).setValue(new Date());
    }
  }

  var currentProgressSync = wmSyncStudentMasterInfoToCurrentProgressBatch_(ss, studentIds);
  wmClearRuntimeCachesForStudents_(studentIds);
  return {success:true, className:className, studentCount:studentIds.length, studentIds:studentIds, currentProgressStudentInfoSync:currentProgressSync};
}

function wmSyncRegisteredStudentToClassDb_(ss, className, studentId, studentName, teacherId, teacherName) {
  className = String(className || '').trim();
  studentId = String(studentId || '').trim();
  studentName = String(studentName || '').trim();
  teacherId = String(teacherId || '').trim();
  teacherName = String(teacherName || '').trim();

  if (!className || !studentId) {
    return {success:true, skipped:true, message:'반 미선택 학생이므로 반관리_DB 갱신을 생략했습니다.'};
  }

  ss = ss || getLmsSpreadsheet_();
  var classSheet = ss.getSheetByName('7-2.반관리_DB');
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  if (!classSheet) return {success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'};
  if (!studentSheet) return {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};

  var classLastRow = classSheet.getLastRow();
  var classLastCol = classSheet.getLastColumn();
  if (classLastRow < 2 || classLastCol < 1) {
    return {success:false, message:'7-2.반관리_DB에 등록된 반이 없습니다.'};
  }

  var classHeaders = classSheet.getRange(1, 1, 1, classLastCol).getDisplayValues()[0].map(function(v){
    return String(v || '').trim();
  });
  var idxClass = classHeaders.indexOf('반명');
  var idxTeacherId = classHeaders.indexOf('담당교사ID');
  var idxTeacherName = classHeaders.indexOf('담당교사명');
  var idxStudentIds = classHeaders.indexOf('학생ID목록');
  var idxStudentNames = classHeaders.indexOf('학생이름목록');
  var idxStudentCount = classHeaders.indexOf('학생수');
  var idxStatus = classHeaders.indexOf('상태');
  var idxUpdatedAt = classHeaders.indexOf('수정일');

  if (idxClass < 0 || idxStudentIds < 0 || idxStudentNames < 0 || idxStudentCount < 0) {
    return {success:false, message:'7-2.반관리_DB 필수 컬럼(반명/학생ID목록/학생이름목록/학생수)을 찾을 수 없습니다.'};
  }

  var classRows = classSheet.getRange(2, 1, classLastRow - 1, classLastCol).getDisplayValues();
  var classKey = className.toUpperCase();
  var targetIndex = -1;
  for (var cr = classRows.length - 1; cr >= 0; cr--) {
    if (String(classRows[cr][idxClass] || '').trim().toUpperCase() === classKey) {
      targetIndex = cr;
      break;
    }
  }
  if (targetIndex < 0) {
    return {success:false, message:'선택한 반을 7-2.반관리_DB에서 찾을 수 없습니다: ' + className};
  }

  var studentLastRow = studentSheet.getLastRow();
  var studentLastCol = studentSheet.getLastColumn();
  var rosterIds = [];
  var rosterNames = [];
  var seen = {};

  if (studentLastRow >= 2 && studentLastCol >= 1) {
    var studentValues = studentSheet.getRange(1, 1, studentLastRow, studentLastCol).getDisplayValues();
    var studentHeaders = studentValues[0].map(function(v){ return String(v || '').trim(); });
    var sIdxId = studentHeaders.indexOf('학생ID');
    var sIdxName = studentHeaders.indexOf('학생이름');
    var sIdxClass = findFirstHeaderIndex_(studentHeaders, ['반명','Class','반','현재반','클래스']);

    if (sIdxId < 0 || sIdxClass < 0) {
      return {success:false, message:'1.학생관리_DB의 학생ID 또는 반명/Class 컬럼을 찾을 수 없습니다.'};
    }

    /* 최신 등록 학생이 위에 보이도록 아래 행부터 역순으로 수집합니다. */
    for (var sr = studentValues.length - 1; sr >= 1; sr--) {
      var studentClassKey = String(studentValues[sr][sIdxClass] || '').trim().toUpperCase();
      if (studentClassKey !== classKey) continue;

      var sid = String(studentValues[sr][sIdxId] || '').trim();
      if (!sid) continue;
      var sidKey = sid.toUpperCase();
      if (seen[sidKey]) continue;
      seen[sidKey] = true;
      rosterIds.push(sid);
      rosterNames.push(String(sIdxName >= 0 ? studentValues[sr][sIdxName] : '').trim() || sid);
    }
  }

  /* 저장 직후 시트 표시값 반영이 지연되는 경우에도 이번 학생은 반드시 포함합니다. */
  var wantedKey = studentId.toUpperCase();
  for (var wi = rosterIds.length - 1; wi >= 0; wi--) {
    if (String(rosterIds[wi] || '').trim().toUpperCase() === wantedKey) {
      rosterIds.splice(wi, 1);
      rosterNames.splice(wi, 1);
    }
  }
  rosterIds.unshift(studentId);
  rosterNames.unshift(studentName || studentId);
  seen[wantedKey] = true;

  var targetRow = targetIndex + 2;
  var outputRow = classRows[targetIndex].slice();
  outputRow[idxStudentIds] = rosterIds.join(', ');
  outputRow[idxStudentNames] = rosterNames.join(', ');
  outputRow[idxStudentCount] = rosterIds.length;
  if (idxTeacherId >= 0 && teacherId) outputRow[idxTeacherId] = teacherId;
  if (idxTeacherName >= 0 && teacherName) outputRow[idxTeacherName] = teacherName;
  if (idxStatus >= 0 && !String(outputRow[idxStatus] || '').trim()) outputRow[idxStatus] = 'Active';
  if (idxUpdatedAt >= 0) outputRow[idxUpdatedAt] = new Date();

  classSheet.getRange(targetRow, 1, 1, classLastCol).setValues([outputRow]);

  return {
    success:true,
    className:String(outputRow[idxClass] || className).trim(),
    studentCount:rosterIds.length,
    studentIds:rosterIds,
    studentNames:rosterNames,
    source:'1.학생관리_DB'
  };
}

function wmMoveStudentClassRosterForLms_(ss, oldClassName, newClassName, studentId, studentName) {
  oldClassName = String(oldClassName || '').trim();
  newClassName = String(newClassName || '').trim();
  studentId = String(studentId || '').trim();
  studentName = String(studentName || studentId).trim();
  if (!studentId || oldClassName === newClassName) return {success:true, skipped:true};
  ss = ss || getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('7-2.반관리_DB');
  if (!sheet) return {success:false, message:'7-2.반관리_DB 시트를 찾을 수 없습니다.'};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {success:false, message:'7-2.반관리_DB에 등록된 반이 없습니다.'};
  var values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function(v){ return String(v || '').trim(); });
  var idxClass = headers.indexOf('반명');
  var idxIds = headers.indexOf('학생ID목록');
  var idxNames = headers.indexOf('학생이름목록');
  var idxCount = headers.indexOf('학생수');
  var idxUpdatedAt = headers.indexOf('수정일');
  if (idxClass < 0 || idxIds < 0 || idxNames < 0 || idxCount < 0) {
    return {success:false, message:'7-2.반관리_DB 필수 컬럼을 확인하세요.'};
  }
  var oldKey = oldClassName.toUpperCase();
  var newKey = newClassName.toUpperCase();
  var oldRow = 0;
  var newRow = 0;
  for (var r = 1; r < values.length; r++) {
    var classKey = String(values[r][idxClass] || '').trim().toUpperCase();
    if (oldKey && classKey === oldKey) oldRow = r + 1;
    if (newKey && classKey === newKey) newRow = r + 1;
  }
  if (newClassName && !newRow) return {success:false, message:'선택한 반을 7-2.반관리_DB에서 찾을 수 없습니다: ' + newClassName};

  function updateRoster_(sheetRow, addStudent) {
    if (!sheetRow) return;
    var row = values[sheetRow - 1].slice();
    var ids = String(row[idxIds] || '').split(',').map(function(v){ return String(v || '').trim(); }).filter(Boolean);
    var names = String(row[idxNames] || '').split(',').map(function(v){ return String(v || '').trim(); });
    var nextIds = [];
    var nextNames = [];
    ids.forEach(function(id, index){
      if (id.toUpperCase() === studentId.toUpperCase()) return;
      nextIds.push(id);
      nextNames.push(names[index] || id);
    });
    if (addStudent) {
      nextIds.unshift(studentId);
      nextNames.unshift(studentName || studentId);
    }
    row[idxIds] = nextIds.join(', ');
    row[idxNames] = nextNames.join(', ');
    row[idxCount] = nextIds.length;
    if (idxUpdatedAt >= 0) row[idxUpdatedAt] = new Date();
    sheet.getRange(sheetRow, 1, 1, lastCol).setValues([row]);
  }

  if (oldRow) updateRoster_(oldRow, false);
  if (newRow) updateRoster_(newRow, true);
  SpreadsheetApp.flush();
  return {success:true, oldClassName:oldClassName, newClassName:newClassName};
}
/* WM_STUDENT_ACTIVE_SERIAL_FAST_V1
 * ACTIVE / INACTIVE 변경 전용 경량 경로
 * 학생 전체행 조회 금지
 * 필요한 학생ID / 활성상태 / 학생일련번호 / 권한확인 셀만 사용
 */

function wmFindStudentRowFast_(sheet, studentId, idIndex) {
  var lastRow = sheet.getLastRow();

  if (!sheet || !studentId || idIndex < 0 || lastRow < 2) {
    return 0;
  }

  var found = sheet
    .getRange(2, idIndex + 1, lastRow - 1, 1)
    .createTextFinder(String(studentId || '').trim())
    .matchEntireCell(true)
    .matchCase(false)
    .findNext();

  return found ? found.getRow() : 0;
}


function wmBuildStudentScopeRowFast_(sheet, headers, targetRow) {
  var result = {};

  var fields = [
    ['교사ID', ['교사ID', '담당교사ID', '현재담당교사ID', 'Teacher_ID', 'teacherId']],
    ['교사명', ['교사명', '담당교사', '현재담당교사', 'Teacher_Name', 'teacher']],
    ['담당리더', ['담당리더', '현재담당리더', '리더', '리더명', 'Leader', 'Leader_Name']],
    ['반명', ['반명', '현재반', 'Class', '반', '담당반']]
  ];

  fields.forEach(function(field) {
    var index = findFirstHeaderIndex_(headers, field[1]);

    if (index >= 0) {
      result[field[0]] = String(
        sheet.getRange(targetRow, index + 1).getDisplayValue() || ''
      ).trim();
    }
  });

  return result;
}


function wmCountActiveStudentsBelowFast_(sheet, activeIndex, targetRow) {
  var lastRow = sheet.getLastRow();
  var startRow = targetRow + 1;

  if (activeIndex < 0 || startRow > lastRow) {
    return 0;
  }

  return sheet
    .getRange(
      startRow,
      activeIndex + 1,
      lastRow - targetRow,
      1
    )
    .createTextFinder('ACTIVE')
    .matchEntireCell(true)
    .matchCase(false)
    .findAll()
    .length;
}


function wmShiftStudentSerialsAboveFast_(
  sheet,
  serialIndex,
  targetRow,
  pivot,
  delta,
  inclusive
) {
  var rowCount = targetRow - 2;

  if (serialIndex < 0 || rowCount <= 0) {
    return { shifted:0 };
  }

  /* 대상 학생보다 위쪽의 학생일련번호 1개 열만 읽습니다. */
  var serialValues = sheet
    .getRange(
      2,
      serialIndex + 1,
      rowCount,
      1
    )
    .getDisplayValues();

  var runs = [];
  var runStart = 0;
  var runValues = [];
  var shifted = 0;

  function closeRun_() {
    if (!runStart || !runValues.length) {
      return;
    }

    runs.push({
      startRow:runStart,
      values:runValues
    });

    runStart = 0;
    runValues = [];
  }

  for (var i = 0; i < serialValues.length; i++) {
    var number = Number(
      String(serialValues[i][0] || '').replace(/[^0-9]/g, '')
    );

    var shouldShift =
      number > 0 &&
      (
        inclusive
          ? number >= pivot
          : number > pivot
      );

    if (shouldShift) {
      if (!runStart) {
        runStart = i + 2;
      }

      runValues.push([number + delta]);
      shifted++;
    } else {
      closeRun_();
    }
  }

  closeRun_();

  /* 실제 번호가 바뀌는 구간만 씁니다. */
  runs.forEach(function(run) {
    sheet
      .getRange(
        run.startRow,
        serialIndex + 1,
        run.values.length,
        1
      )
      .setValues(run.values);
  });

  return {
    shifted:shifted
  };
}


function wmSaveStudentActiveStatusFast_(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    /* 기다리지 않습니다. 다른 상태변경과 겹치면 즉시 종료합니다. */
    locked = lock.tryLock(0);

    if (!locked) {
      return {
        success:false,
        message:'다른 학생 상태변경이 처리 중입니다. 다시 저장하세요.'
      };
    }

    var p = (e && e.parameter) ? e.parameter : {};

    var studentId = String(
      p.studentId ||
      p.studentID ||
      p.sid ||
      ''
    ).trim().toUpperCase();

    var nextStatus = String(
      p.activeStatus ||
      p['활성상태'] ||
      ''
    ).trim().toUpperCase();

    if (!studentId) {
      return {
        success:false,
        message:'학생ID가 없습니다.'
      };
    }

    if (['ACTIVE', 'INACTIVE'].indexOf(nextStatus) === -1) {
      return {
        success:false,
        message:'활성상태는 ACTIVE 또는 INACTIVE만 선택할 수 있습니다.'
      };
    }

    var actor = wmGetLmsRequestActor_(e);

    if (
      String(p.actorTeacherId || '').trim() &&
      !actor.found
    ) {
      return {
        success:false,
        message:'로그인 교사 정보를 확인할 수 없습니다.'
      };
    }

    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('1.학생관리_DB');

    if (!sheet) {
      return {
        success:false,
        message:'1.학생관리_DB 시트를 찾을 수 없습니다.'
      };
    }

    /* 데이터 전체가 아니라 헤더 1줄만 읽습니다. */
    var lastColumn = sheet.getLastColumn();

    var headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(function(value) {
        return String(value || '').trim();
      });

    var idxStudentId = headers.indexOf('학생ID');
    var idxActiveStatus = headers.indexOf('활성상태');
    var idxSerial = headers.indexOf('학생일련번호');

    if (
      idxStudentId < 0 ||
      idxActiveStatus < 0 ||
      idxSerial < 0
    ) {
      return {
        success:false,
        message:'학생ID·활성상태·학생일련번호 컬럼을 확인하세요.'
      };
    }

    /* 학생ID 열에서 요청한 학생 한 명만 찾습니다. */
    var targetRow = wmFindStudentRowFast_(
      sheet,
      studentId,
      idxStudentId
    );

    if (!targetRow) {
      return {
        success:false,
        message:'학생ID를 찾을 수 없습니다: ' + studentId
      };
    }

    /* SUPER / SPECIAL ADMIN은 불필요한 학생 권한셀도 읽지 않습니다. */
    if (actor.found) {
      var actorRole = wmNormalizeLmsRole_(actor.role);

      if (
        actorRole === 'TEACHER' ||
        actorRole === 'SUB_LEADER'
      ) {
        var scopeRow = wmBuildStudentScopeRowFast_(
          sheet,
          headers,
          targetRow
        );

        if (!wmStudentBelongsToActor_(scopeRow, actor)) {
          return {
            success:false,
            message:'현재 권한으로 수정할 수 없는 학생입니다.'
          };
        }
      }
    }

    /* 해당 학생의 필요한 셀 2개만 읽습니다. */
    var currentStatus = String(
      sheet
        .getRange(targetRow, idxActiveStatus + 1)
        .getDisplayValue() || ''
    ).trim().toUpperCase();

    var currentSerialText = String(
      sheet
        .getRange(targetRow, idxSerial + 1)
        .getDisplayValue() || ''
    ).trim();

    if (
      ['ACTIVE', 'INACTIVE'].indexOf(currentStatus) === -1
    ) {
      return {
        success:false,
        message:'현재 활성상태 값을 확인하세요: ' + currentStatus
      };
    }

    /* 이미 같은 상태면 아무것도 다시 쓰지 않습니다. */
    if (currentStatus === nextStatus) {
      return {
        success:true,
        unchanged:true,
        studentId:studentId,
        activeStatus:currentStatus,
        studentSerial:currentSerialText,
        currentProgressStudentInfoSync:{
          success:true,
          skipped:true
        }
      };
    }


    /* =====================================================
       ACTIVE → INACTIVE
       해당 학생 번호 제거
       필요한 위쪽 학생일련번호만 -1
       ===================================================== */
    if (
      currentStatus === 'ACTIVE' &&
      nextStatus === 'INACTIVE'
    ) {
      var removedSerial = Number(
        currentSerialText.replace(/[^0-9]/g, '')
      );

      if (!(removedSerial > 0)) {
        return {
          success:false,
          message:'현재 학생일련번호를 먼저 정상화해야 합니다.'
        };
      }

      var downShift = wmShiftStudentSerialsAboveFast_(
        sheet,
        idxSerial,
        targetRow,
        removedSerial,
        -1,
        false
      );

      sheet
        .getRange(targetRow, idxActiveStatus + 1)
        .setValue('INACTIVE');

      sheet
        .getRange(targetRow, idxSerial + 1)
        .clearContent();

              wmClearRuntimeCachesForStudent_(studentId);
              wmMarkDashboardDataChanged_();

      return {
        success:true,
        studentId:studentId,
        activeStatus:'INACTIVE',
        studentSerial:'',
        serialShift:{
          type:'DOWN_AFTER',
          pivot:removedSerial,
          shifted:downShift.shifted
        },
        currentProgressStudentInfoSync:{
          success:true,
          skipped:true
        }
      };
    }


    /* =====================================================
       INACTIVE → ACTIVE
       아래쪽 ACTIVE 수만 확인하여 자기 번호 계산
       필요한 위쪽 학생일련번호만 +1
       ===================================================== */
    var newSerial = 1;
    var serialLastRow = sheet.getLastRow();

    if (targetRow < serialLastRow) {
      var belowSerialCell = sheet
        .getRange(
          targetRow + 1,
          idxSerial + 1,
          serialLastRow - targetRow,
          1
        )
        .createTextFinder('^[0-9]+$')
        .useRegularExpression(true)
        .matchEntireCell(true)
        .findNext();

      if (belowSerialCell) {
        newSerial =
          Number(belowSerialCell.getDisplayValue() || 0) + 1;
      }
    }

    var upShift = wmShiftStudentSerialsAboveFast_(
      sheet,
      idxSerial,
      targetRow,
      newSerial,
      1,
      true
    );

    sheet
      .getRange(targetRow, idxActiveStatus + 1)
      .setValue('ACTIVE');

    sheet
      .getRange(targetRow, idxSerial + 1)
      .setValue(newSerial);

    wmClearRuntimeCachesForStudent_(studentId);
        wmMarkDashboardDataChanged_();
        
    return {
      success:true,
      studentId:studentId,
      activeStatus:'ACTIVE',
      studentSerial:String(newSerial),
      serialShift:{
        type:'UP_FROM',
        pivot:newSerial,
        shifted:upShift.shifted
      },
      currentProgressStudentInfoSync:{
        success:true,
        skipped:true
      }
    };

  } catch (err) {
    return {
      success:false,
      message:'활성상태 변경 오류',
      error:String(
        err && err.message
          ? err.message
          : err
      )
    };

  } finally {
    if (locked) {
      try {
        lock.releaseLock();
      } catch (ignore) {}
    }
  }
}
/* WM_STUDENT_DELETE_FAST_HELPERS_V1
 * DELETE 전용 경량 helper
 * 전체행 조회 금지
 * 학생ID 열 / 해당 반의 명단 셀만 사용
 */

function wmFindStudentRowsByIdFast_(sheet, studentId, idIndex) {
  var lastRow = sheet.getLastRow();

  if (
    !sheet ||
    !studentId ||
    idIndex < 0 ||
    lastRow < 2
  ) {
    return [];
  }

  return sheet
    .getRange(
      2,
      idIndex + 1,
      lastRow - 1,
      1
    )
    .createTextFinder(
      String(studentId || '').trim()
    )
    .matchEntireCell(true)
    .matchCase(false)
    .findAll()
    .map(function(cell) {
      return cell.getRow();
    });
}


function wmDeleteRowsGroupedFast_(sheet, rowNumbers) {
  var rows = (rowNumbers || [])
    .map(function(row) {
      return Number(row || 0);
    })
    .filter(function(row) {
      return row >= 2;
    })
    .sort(function(a, b) {
      return b - a;
    });

  if (!rows.length) {
    return 0;
  }

  var uniqueRows = [];
  var seen = {};

  rows.forEach(function(row) {
    if (seen[row]) {
      return;
    }

    seen[row] = true;
    uniqueRows.push(row);
  });

  var groups = [];
  var high = uniqueRows[0];
  var low = high;

  for (var i = 1; i < uniqueRows.length; i++) {
    var row = uniqueRows[i];

    if (row === low - 1) {
      low = row;
      continue;
    }

    groups.push({
      start:low,
      count:high - low + 1
    });

    high = row;
    low = row;
  }

  groups.push({
    start:low,
    count:high - low + 1
  });

  groups.forEach(function(group) {
    sheet.deleteRows(
      group.start,
      group.count
    );
  });

  return uniqueRows.length;
}


function wmRemoveStudentFromClassRosterDeleteFast_(
  ss,
  className,
  studentId
) {
  className = String(className || '').trim();
  studentId = String(studentId || '').trim();

  if (!studentId) {
    return {
      success:true,
      skipped:true
    };
  }

  var sheet = ss.getSheetByName(
    '7-2.반관리_DB'
  );

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {
    return {
      success:true,
      skipped:true
    };
  }

  /* 헤더 1줄만 읽습니다. */
  var headers = sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  var idxClass =
    headers.indexOf('반명');

  var idxIds =
    headers.indexOf('학생ID목록');

  var idxNames =
    headers.indexOf('학생이름목록');

  var idxCount =
    headers.indexOf('학생수');

  var idxUpdated =
    headers.indexOf('수정일');

  if (
    idxClass < 0 ||
    idxIds < 0 ||
    idxNames < 0 ||
    idxCount < 0
  ) {
    return {
      success:false,
      message:'7-2.반관리_DB 필수 컬럼을 확인하세요.'
    };
  }

  /* 반명으로 찾지 않고 학생ID목록에서
     삭제 대상 학생ID가 실제 들어 있는 반만 찾습니다. */
  var escapedStudentId = studentId.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );

  var studentListCell = sheet
    .getRange(
      2,
      idxIds + 1,
      sheet.getLastRow() - 1,
      1
    )
    .createTextFinder(
      '(^|,\\s*)' +
      escapedStudentId +
      '(\\s*,|$)'
    )
    .useRegularExpression(true)
    .matchCase(false)
    .findNext();

  if (!studentListCell) {
    return {
      success:true,
      skipped:true
    };
  }

  var targetRow = studentListCell.getRow();

  /* 해당 반의 학생ID목록 / 학생이름목록 2셀만 읽습니다. */
  var ids = String(
    sheet
      .getRange(
        targetRow,
        idxIds + 1
      )
      .getDisplayValue() || ''
  )
    .split(',')
    .map(function(value) {
      return String(value || '').trim();
    })
    .filter(Boolean);

  var names = String(
    sheet
      .getRange(
        targetRow,
        idxNames + 1
      )
      .getDisplayValue() || ''
  )
    .split(',')
    .map(function(value) {
      return String(value || '').trim();
    });

  var wanted =
    studentId.toUpperCase();

  var nextIds = [];
  var nextNames = [];

  ids.forEach(function(id, index) {
    if (
      String(id || '')
        .trim()
        .toUpperCase() === wanted
    ) {
      return;
    }

    nextIds.push(id);
    nextNames.push(
      names[index] || id
    );
  });

  if (nextIds.length === ids.length) {
    return {
      success:true,
      skipped:true
    };
  }

  /* 바뀌는 셀만 씁니다. */
  sheet
    .getRange(
      targetRow,
      idxIds + 1
    )
    .setValue(
      nextIds.join(', ')
    );

  sheet
    .getRange(
      targetRow,
      idxNames + 1
    )
    .setValue(
      nextNames.join(', ')
    );

  sheet
    .getRange(
      targetRow,
      idxCount + 1
    )
    .setValue(
      nextIds.length
    );

  if (idxUpdated >= 0) {
    sheet
      .getRange(
        targetRow,
        idxUpdated + 1
      )
      .setValue(
        new Date()
      );
  }

  return {
    success:true,
    removed:true,
    studentCount:nextIds.length
  };
}

/* WM_DELETED_STUDENT_ARCHIVE_V1
 * DELETE 직전 1.학생관리_DB 해당 학생 1행을
 * 1-1.삭제학생_DB에 먼저 보관합니다.
 * 보관 실패 시 실제 DELETE를 진행하지 않습니다.
 */

/* WM_DELETED_STUDENT_LMS_READ_V1
 * SUPER_ADMIN / SPECIAL_ADMIN이 요청할 때만
 * 1-1.삭제학생_DB를 조회합니다.
 * 평상시 LMS 로딩에는 연결하지 않습니다.
 */
function wmGetDeletedStudentsForLms_(e) {
  try {
    var actor =
      wmGetLmsRequestActor_(e);

    if (!actor || !actor.found) {
      return {
        success:false,
        message:
          '로그인 교사 정보를 확인할 수 없습니다.'
      };
    }

    var role =
      wmNormalizeLmsRole_(
        actor.role || ''
      );

    if (
      role !== 'SUPER_ADMIN' &&
      role !== 'SPECIAL_ADMIN'
    ) {
      return {
        success:false,
        message:
          '삭제학생 조회는 수퍼어드민 또는 스페셜어드민만 가능합니다.'
      };
    }

    var ss =
      getLmsSpreadsheet_();

    var sheet =
      ss.getSheetByName(
        '1-1.삭제학생_DB'
      );

    if (!sheet) {
      return {
        success:false,
        message:
          '1-1.삭제학생_DB 시트를 찾을 수 없습니다.'
      };
    }

    var lastRow =
      sheet.getLastRow();

    var lastColumn =
      sheet.getLastColumn();

    if (lastColumn < 1) {
      return {
        success:true,
        headers:[],
        rows:[]
      };
    }

    var headers =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getDisplayValues()[0]
        .map(function(value) {
          return String(
            value || ''
          ).trim();
        });

    if (lastRow < 2) {
      return {
        success:true,
        sheetName:sheet.getName(),
        headers:headers,
        rows:[]
      };
    }

    /*
     * 삭제학생 조회 버튼을 눌렀을 때만
     * 삭제학생_DB 자체를 읽습니다.
     */
    var values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          lastColumn
        )
        .getDisplayValues();

    var rows = [];

    /*
     * 최근 DELETE가 위에 보이도록
     * 마지막 행부터 역순으로 반환합니다.
     */
    for (
      var r = values.length - 1;
      r >= 0;
      r--
    ) {
      var hasValue =
        values[r].some(
          function(value) {
            return String(
              value || ''
            ).trim() !== '';
          }
        );

      if (!hasValue) {
        continue;
      }

      var rowObject = {};

      headers.forEach(
        function(header, index) {
          if (!header) {
            return;
          }

          rowObject[header] =
            String(
              values[r][index] || ''
            ).trim();
        }
      );

      rows.push(rowObject);
    }

    return {
      success:true,
      sheetName:sheet.getName(),
      headers:headers,
      rows:rows
    };

  } catch (err) {
    return {
      success:false,
      message:
        '삭제학생_DB 조회 오류',
      error:String(
        err && err.message
          ? err.message
          : err
      )
    };
  }
}

function wmArchiveDeletedStudentForLms_(
  ss,
  studentHeaders,
  originalRow,
  actor
) {
  var archiveSheet = ss.getSheetByName(
    '1-1.삭제학생_DB'
  );

  if (!archiveSheet) {
    return {
      success:false,
      message:'1-1.삭제학생_DB 시트를 찾을 수 없습니다.'
    };
  }

  var archiveLastColumn =
    archiveSheet.getLastColumn();

  if (archiveLastColumn < 1) {
    return {
      success:false,
      message:'1-1.삭제학생_DB 헤더를 확인하세요.'
    };
  }

  var archiveHeaders = archiveSheet
    .getRange(
      1,
      1,
      1,
      archiveLastColumn
    )
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  var requiredHeaders = [
    '삭제일시',
    '삭제관리자ID',
    '학생ID',
    '학생이름',
    '학생일련번호',
    '활성상태',
    '반명',
    '교사ID',
    '교사명',
    '학교',
    '학년',
    '등록일',
    '최초등록일',
    '최근결제일',
    '비고',
    '원본학생관리데이터'
  ];

  var missingHeaders = requiredHeaders
    .filter(function(header) {
      return archiveHeaders.indexOf(header) < 0;
    });

  if (missingHeaders.length) {
    return {
      success:false,
      message:
        '1-1.삭제학생_DB 필수 컬럼 누락: ' +
        missingHeaders.join(', ')
    };
  }

  var original = {};

  studentHeaders.forEach(function(header, index) {
    if (!header) {
      return;
    }

    original[header] = String(
      originalRow[index] || ''
    ).trim();
  });

  function pick_(names) {
    for (var i = 0; i < names.length; i++) {
      var key = names[i];

      if (
        Object.prototype.hasOwnProperty.call(
          original,
          key
        )
      ) {
        return String(
          original[key] || ''
        ).trim();
      }
    }

    return '';
  }

  var archiveObject = {
    '삭제일시':new Date(),

    '삭제관리자ID':String(
      actor && actor.teacherId || ''
    ).trim(),

    '학생ID':pick_([
      '학생ID'
    ]),

    '학생이름':pick_([
      '학생이름',
      '학생명',
      '이름'
    ]),

    '학생일련번호':pick_([
      '학생일련번호'
    ]),

    '활성상태':pick_([
      '활성상태'
    ]),

    '반명':pick_([
      '반명',
      'Class',
      '반'
    ]),

    '교사ID':pick_([
      '교사ID',
      '담당교사ID',
      '현재담당교사ID'
    ]),

    '교사명':pick_([
      '교사명',
      '담당교사',
      '현재담당교사'
    ]),

    '학교':pick_([
      '학교'
    ]),

    '학년':pick_([
      '학년'
    ]),

    '등록일':pick_([
      '등록일'
    ]),

    '최초등록일':pick_([
      '최초등록일'
    ]),

    '최근결제일':pick_([
      '최근결제일'
    ]),

    '비고':pick_([
      '비고',
      '메모'
    ]),

    '원본학생관리데이터':
      JSON.stringify(original)
  };

  var archiveRow = Math.max(
    2,
    archiveSheet.getLastRow() + 1
  );

  var outputRow = archiveHeaders.map(
    function(header) {
      return Object.prototype.hasOwnProperty.call(
        archiveObject,
        header
      )
        ? archiveObject[header]
        : '';
    }
  );

  archiveSheet
    .getRange(
      archiveRow,
      1,
      1,
      archiveHeaders.length
    )
    .setValues([outputRow]);

  /* 실제 보관 완료를 확인한 뒤에만 DELETE 허용 */
  SpreadsheetApp.flush();

  var archiveStudentIdIndex =
    archiveHeaders.indexOf('학생ID');

  var savedStudentId = String(
    archiveSheet
      .getRange(
        archiveRow,
        archiveStudentIdIndex + 1
      )
      .getDisplayValue() || ''
  ).trim().toUpperCase();

  var expectedStudentId = String(
    archiveObject['학생ID'] || ''
  ).trim().toUpperCase();

  if (
    !expectedStudentId ||
    savedStudentId !== expectedStudentId
  ) {
    return {
      success:false,
      message:'삭제학생_DB 보관 확인에 실패했습니다.'
    };
  }

  return {
    success:true,
    archiveRow:archiveRow,
    studentId:expectedStudentId
  };
}

function wmDeleteStudentForLms_(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    locked = lock.tryLock(0);

    if (!locked) {
      return {
        success:false,
        message:'다른 학생 작업이 처리 중입니다. 다시 시도하세요.'
      };
    }

    var p = (e && e.parameter) ? e.parameter : {};

    var studentId = String(
      p.studentId ||
      p.studentID ||
      p['학생ID'] ||
      ''
    ).trim().toUpperCase();

    if (!studentId) {
      return {
        success:false,
        message:'삭제할 학생ID가 없습니다.'
      };
    }

    var actor = wmGetLmsRequestActor_(e);

    if (!actor || !actor.found) {
      return {
        success:false,
        message:'로그인 교사 정보를 확인할 수 없습니다.'
      };
    }

    var actorRole = wmNormalizeLmsRole_(
      actor.role || ''
    );

    if (
      actorRole !== 'SUPER_ADMIN' &&
      actorRole !== 'SPECIAL_ADMIN'
    ) {
      return {
        success:false,
        message:'학생 DELETE는 수퍼어드민 또는 스페셜어드민만 가능합니다.'
      };
    }

    var ss = getLmsSpreadsheet_();
    var studentSheet = ss.getSheetByName(
      '1.학생관리_DB'
    );

    if (!studentSheet) {
      return {
        success:false,
        message:'1.학생관리_DB 시트를 찾을 수 없습니다.'
      };
    }

    var studentLastColumn =
      studentSheet.getLastColumn();

    var studentHeaders = studentSheet
      .getRange(
        1,
        1,
        1,
        studentLastColumn
      )
      .getDisplayValues()[0]
      .map(function(value) {
        return String(value || '').trim();
      });

    var idxStudentId =
      studentHeaders.indexOf('학생ID');

    var idxStudentName =
      studentHeaders.indexOf('학생이름');

    var idxActiveStatus =
      studentHeaders.indexOf('활성상태');

    var idxClassName =
      findFirstHeaderIndex_(
        studentHeaders,
        ['반명', 'Class', '반']
      );

    if (
      idxStudentId < 0 ||
      idxActiveStatus < 0
    ) {
      return {
        success:false,
        message:'학생ID·활성상태 컬럼을 확인하세요.'
      };
    }

    /* 학생ID 열에서 삭제 대상 학생 1명만 찾습니다. */
    var studentRow = wmFindStudentRowFast_(
      studentSheet,
      studentId,
      idxStudentId
    );

    if (!studentRow) {
      return {
        success:false,
        message:'삭제할 학생을 찾을 수 없습니다: ' + studentId
      };
    }

    /* 안전장치:
       해당 학생 활성상태 1셀만 확인합니다. */
    var currentStatus = String(
      studentSheet
        .getRange(
          studentRow,
          idxActiveStatus + 1
        )
        .getDisplayValue() || ''
    ).trim().toUpperCase();

    if (currentStatus !== 'INACTIVE') {
      return {
        success:false,
        message:
          'ACTIVE 상태는 DELETE가 되지 않습니다.\n' +
          'DELETE하려면 해당 학생을 먼저 INACTIVE 하세요.'
      };
    }

    /*
     * DELETE 보관용으로
     * 해당 학생 1행만 1회 읽습니다.
     * 다른 학생 행은 읽지 않습니다.
     */
    var originalStudentRow = studentSheet
      .getRange(
        studentRow,
        1,
        1,
        studentLastColumn
      )
      .getDisplayValues()[0];

    var studentName =
      idxStudentName >= 0
        ? String(
            originalStudentRow[idxStudentName] || ''
          ).trim()
        : '';

    var className =
      idxClassName >= 0
        ? String(
            originalStudentRow[idxClassName] || ''
          ).trim()
        : '';

    /*
     * 실제 학생ID가 있는 DB만 대상.
     * 삭제 시작 전에 모든 대상 시트의
     * 학생ID 헤더부터 먼저 확인합니다.
     */
    var targetSheetNames = [
      '2.학습기록_DB',
      '4-1.성적표생성_DB',
      '5-2.자동발송설정_DB',
      '6.성적등급_DB',
      '8.현재진행_DB'
    ];

    var targets = [];

    for (
      var i = 0;
      i < targetSheetNames.length;
      i++
    ) {
      var sheetName =
        targetSheetNames[i];

      var targetSheet =
        ss.getSheetByName(sheetName);

      if (!targetSheet) {
        return {
          success:false,
          message:
            sheetName +
            ' 시트를 찾을 수 없습니다.'
        };
      }

      var targetLastColumn =
        targetSheet.getLastColumn();

      if (targetLastColumn < 1) {
        return {
          success:false,
          message:
            sheetName +
            ' 헤더를 확인하세요.'
        };
      }

      var targetHeaders = targetSheet
        .getRange(
          1,
          1,
          1,
          targetLastColumn
        )
        .getDisplayValues()[0]
        .map(function(value) {
          return String(value || '').trim();
        });

      var targetIdIndex =
        wmFindHeaderIndex_(
          targetHeaders,
          [
            '학생ID',
            'Student_ID',
            'studentId'
          ]
        );

      if (targetIdIndex < 0) {
        return {
          success:false,
          message:
            sheetName +
            ' 학생ID 컬럼을 확인하세요.'
        };
      }

      targets.push({
        name:sheetName,
        sheet:targetSheet,
        idIndex:targetIdIndex
      });
    }

    /*
     * 실제 삭제 전에
     * 1-1.삭제학생_DB에 먼저 보관합니다.
     *
     * 보관 실패 시 여기서 즉시 종료하므로
     * 아래 운영 DB DELETE는 시작되지 않습니다.
     */
    /*
     * INACTIVE 처리하면서 학생일련번호는 이미 비워졌으므로
     * DELETE 직전 행 위치에서 기존 학생일련번호만 재구성합니다.
     *
     * 1.학생관리_DB에는 쓰지 않습니다.
     * 학생일련번호 열에서 아래쪽 첫 정상번호 1개만 찾습니다.
     */
    var archiveStudentRow =
      originalStudentRow.slice();

    var archiveSerialIndex =
      studentHeaders.indexOf(
        '학생일련번호'
      );

    if (archiveSerialIndex >= 0) {
      var deletedStudentSerial = 1;
      var archiveLastRow =
        studentSheet.getLastRow();

      if (studentRow < archiveLastRow) {
        var belowSerialCell =
          studentSheet
            .getRange(
              studentRow + 1,
              archiveSerialIndex + 1,
              archiveLastRow - studentRow,
              1
            )
            .createTextFinder('^[0-9]+$')
            .useRegularExpression(true)
            .matchEntireCell(true)
            .findNext();

        if (belowSerialCell) {
          deletedStudentSerial =
            Number(
              belowSerialCell
                .getDisplayValue() || 0
            ) + 1;
        }
      }

      archiveStudentRow[
        archiveSerialIndex
      ] = String(
        deletedStudentSerial
      );
    }

    var archiveResult =
      wmArchiveDeletedStudentForLms_(
        ss,
        studentHeaders,
        archiveStudentRow,
        actor
      );

    if (!archiveResult.success) {
      return archiveResult;
    }

    /*
     * 반관리_DB:
     * 전체 학생목록을 다시 만들지 않고
     * 해당 반의 명단 셀에서 이 학생만 제거합니다.
     */
    var classRosterResult =
      wmRemoveStudentFromClassRosterDeleteFast_(
        ss,
        className,
        studentId
      );

    if (!classRosterResult.success) {
      return classRosterResult;
    }

    var deletedBySheet = {};

    /*
     * 각 DB마다 학생ID 열 1개에서
     * 해당 학생ID만 찾습니다.
     */
    targets.forEach(function(target) {
      var rowNumbers =
        wmFindStudentRowsByIdFast_(
          target.sheet,
          studentId,
          target.idIndex
        );

      deletedBySheet[target.name] =
        wmDeleteRowsGroupedFast_(
          target.sheet,
          rowNumbers
        );
    });

    /*
     * 연결 DB 삭제가 끝난 뒤
     * 마지막으로 학생관리_DB의
     * 해당 학생 1행만 삭제합니다.
     */
    studentSheet.deleteRow(
      studentRow
    );

    wmClearRuntimeCachesForStudent_(
      studentId
    );

    return {
      success:true,
      message:'학생 DELETE 완료',
      studentId:studentId,
      studentName:studentName,
      deletedBySheet:deletedBySheet,
      classRoster:classRosterResult
    };

  } catch (err) {
    return {
      success:false,
      message:'학생 DELETE 오류',
      error:String(
        err && err.message
          ? err.message
          : err
      )
    };

  } finally {
    if (locked) {
      try {
        lock.releaseLock();
      } catch (ignore) {}
    }
  }
}

/* WM_ACTIVE_STUDENT_SERIAL_NORMALIZE_ONCE_V1
 * 기존 1111 / 빈칸 / 잘못된 학생일련번호를 최초 1회만 정상화합니다.
 * 전체 학생행은 읽지 않습니다.
 * 읽기: 활성상태 1열 + 학생일련번호 1열
 * 쓰기: 실제 틀린 학생일련번호 셀만
 */
function wmNormalizeActiveStudentSerialsOnce() {
  var sheet = getLmsSpreadsheet_().getSheetByName('1.학생관리_DB');

  if (!sheet) {
    return {
      success:false,
      message:'1.학생관리_DB 시트를 찾을 수 없습니다.'
    };
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return {
      success:true,
      activeCount:0,
      updatedCells:0
    };
  }

  var headers = sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  var idxActiveStatus =
    headers.indexOf('활성상태');

  var idxStudentSerial =
    headers.indexOf('학생일련번호');

  if (
    idxActiveStatus < 0 ||
    idxStudentSerial < 0
  ) {
    return {
      success:false,
      message:'활성상태·학생일련번호 컬럼을 확인하세요.'
    };
  }

  var rowCount = lastRow - 1;

  /* 필요한 두 열만 읽습니다. */
  var activeValues = sheet
    .getRange(
      2,
      idxActiveStatus + 1,
      rowCount,
      1
    )
    .getDisplayValues();

  var serialValues = sheet
    .getRange(
      2,
      idxStudentSerial + 1,
      rowCount,
      1
    )
    .getDisplayValues();

  var expected = new Array(rowCount);
  var activeCount = 0;

  /*
   * 신규학생은 현재 2행에 들어오므로
   * 오래된 ACTIVE부터 1,
   * 최신 ACTIVE가 N이 되게 만듭니다.
   */
  for (var i = rowCount - 1; i >= 0; i--) {
    var status = String(
      activeValues[i][0] || ''
    ).trim().toUpperCase();

    if (status === 'ACTIVE') {
      activeCount++;
      expected[i] = String(activeCount);
    } else {
      expected[i] = '';
    }
  }

  var runs = [];
  var runStart = 0;
  var runValues = [];
  var updatedCells = 0;

  function closeRun_() {
    if (!runStart || !runValues.length) {
      return;
    }

    runs.push({
      startRow:runStart,
      values:runValues
    });

    runStart = 0;
    runValues = [];
  }

  for (var r = 0; r < rowCount; r++) {
    var current = String(
      serialValues[r][0] || ''
    ).trim();

    if (current !== expected[r]) {
      if (!runStart) {
        runStart = r + 2;
      }

      runValues.push([
        expected[r] === ''
          ? ''
          : Number(expected[r])
      ]);

      updatedCells++;
    } else {
      closeRun_();
    }
  }

  closeRun_();

  /* 정상 셀은 쓰지 않고 실제 틀린 구간만 수정합니다. */
  runs.forEach(function(run) {
    sheet
      .getRange(
        run.startRow,
        idxStudentSerial + 1,
        run.values.length,
        1
      )
      .setValues(run.values);
  });

  return {
    success:true,
    activeCount:activeCount,
    updatedCells:updatedCells
  };
}

/* 학생일련번호 직접 오입력 시 편집한 셀만 복구 */
function wmRepairEditedStudentSerialFast_(sheet, headers, rowNumber) {
  if (!sheet || rowNumber < 2) {
    return {
      success:true,
      skipped:true
    };
  }

  var idxActiveStatus =
    headers.indexOf('활성상태');

  var idxStudentSerial =
    headers.indexOf('학생일련번호');

  if (
    idxActiveStatus < 0 ||
    idxStudentSerial < 0
  ) {
    return {
      success:false,
      message:'활성상태·학생일련번호 컬럼을 확인하세요.'
    };
  }

  var status = String(
    sheet
      .getRange(
        rowNumber,
        idxActiveStatus + 1
      )
      .getDisplayValue() || ''
  ).trim().toUpperCase();

  var expectedSerial = '';

  if (status === 'ACTIVE') {
    var serialLastRow = sheet.getLastRow();
    var expectedNumber = 1;

    if (rowNumber < serialLastRow) {
      var belowSerialCell = sheet
        .getRange(
          rowNumber + 1,
          idxStudentSerial + 1,
          serialLastRow - rowNumber,
          1
        )
        .createTextFinder('^[0-9]+$')
        .useRegularExpression(true)
        .matchEntireCell(true)
        .findNext();

      if (belowSerialCell) {
        expectedNumber =
          Number(belowSerialCell.getDisplayValue() || 0) + 1;
      }
    }

    expectedSerial = String(expectedNumber);
  }

  var currentSerial = String(
    sheet
      .getRange(
        rowNumber,
        idxStudentSerial + 1
      )
      .getDisplayValue() || ''
  ).trim();

  if (currentSerial === expectedSerial) {
    return {
      success:true,
      skipped:true
    };
  }

  var serialCell = sheet.getRange(
    rowNumber,
    idxStudentSerial + 1
  );

  if (expectedSerial) {
    serialCell.setValue(
      Number(expectedSerial)
    );
  } else {
    serialCell.clearContent();
  }

  return {
    success:true,
    repaired:true,
    studentSerial:expectedSerial
  };
}

function wmSaveStudentToDbForLms_(e) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(2000)) return {success:false, message:'학생등록이 동시에 처리 중입니다. 잠시 후 다시 저장하세요.'};
    var p = (e && e.parameter) ? e.parameter : {};
    var actor = wmGetLmsRequestActor_(e);
    if (String(p.actorTeacherId || '').trim() && !actor.found) {
      return {success:false, message:'로그인 교사 정보를 확인할 수 없습니다.'};
    }
    var actorAccess = wmStudentDataAccessForActor_(actor);
    var requestedParentData = String(p.recipientType || p['수신자구분'] || p['수신구분'] || p['수신'] || p.parentName || p['학부모명'] || p.parentPhone || p['학부모연락처'] || '').trim();
    var requestedPaymentData = String(
      p.lastPayment || p['최종결제'] || p.registerDate || p['등록일'] || p.firstRegisterDate || p['최초등록일'] ||
      p.recentPaymentDate || p['최근결제일'] || p.paymentGapDays || p['결제공백일수'] ||
      p.reregisterCount || p['재등록횟수'] || p.totalRegisterPeriod || p['총등록기간'] || ''
    ).trim();
    if (actor.found && requestedParentData && !actorAccess.parent) return {success:false, message:'학부모정보 입력권한이 없습니다.'};
    if (actor.found && requestedPaymentData && !actorAccess.payment) return {success:false, message:'결제정보 입력권한이 없습니다.'};
    var required = [
      ['studentId','학생ID'], ['studentName','학생이름'], ['password','비밀번호'],
      ['attendancePromise','출석약속'], ['setPromise','세트약속'],
      ['learningAssign','최초배정'], ['s4','S4실루엣단계'], ['s6','S6테스트모드'],
      ['levelCompleteCondition','레벨완료조건']
    ];
    for (var r = 0; r < required.length; r++) {
      if (!String(p[required[r][0]] || p[required[r][1]] || '').trim()) {
        return {success:false, message:required[r][1] + ' 항목을 입력 또는 선택하세요.'};
      }
    }
    var studentId = String(p.studentId || p['학생ID'] || '').trim();
    if (!/^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]{4,10}$/.test(studentId)) {
      return {success:false, message:'학생ID는 영문+숫자 4~10자입니다.'};
    }
    var password = String(p.password || p['비밀번호'] || '').trim();
    if (!/^\d{4}$/.test(password)) return {success:false, message:'비밀번호는 숫자 4자리로 입력하세요.'};

    var registerDateRaw = String(p.registerDate || p['등록일'] || '').trim();
    var learningStartDate = wmNormalizeStudentLearningDate_(p.learningStartDate || p['학습시작일'] || registerDateRaw);
    var learningEndDate = wmNormalizeStudentLearningDate_(p.learningEndDate || p['학습종료일'] || '');
    if (learningStartDate === null) return {success:false, message:'학습시작일 형식이 올바르지 않습니다.'};
    if (!learningStartDate) return {success:false, message:'학습시작일을 달력에서 선택하세요.'};
    if (learningEndDate === null) return {success:false, message:'학습종료일 형식이 올바르지 않습니다.'};
    if (learningEndDate && learningEndDate < learningStartDate) {
      return {success:false, message:'학습종료일은 학습시작일보다 빠를 수 없습니다.'};
    }

    var duplicate = wmCheckStudentIdForLms_({parameter:{studentId:studentId}});
    if (!duplicate.success || !duplicate.available) return duplicate;

    var sheet = getLmsSpreadsheet_().getSheetByName('1.학생관리_DB');
    if (!sheet) return {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
        var serialIndex = headers.indexOf('학생일련번호');
    var activeIndex = headers.indexOf('활성상태');

    if (serialIndex < 0 || activeIndex < 0) {
      return {
        success:false,
        message:'학생일련번호·활성상태 컬럼을 확인하세요.'
      };
    }

    var newActiveStatus = String(
      p.activeStatus ||
      p['활성상태'] ||
      'ACTIVE'
    ).trim().toUpperCase();

    if (['ACTIVE', 'INACTIVE'].indexOf(newActiveStatus) === -1) {
      return {
        success:false,
        message:'신규학생 활성상태는 ACTIVE 또는 INACTIVE만 선택할 수 있습니다.'
      };
    }

    var studentSerial = '';

    if (newActiveStatus === 'ACTIVE') {
      var serialLastRow = sheet.getLastRow();
      var currentActiveCount = 0;

      if (serialLastRow >= 2) {
        var topSerialCell = sheet
          .getRange(
            2,
            serialIndex + 1,
            serialLastRow - 1,
            1
          )
          .createTextFinder('^[0-9]+$')
          .useRegularExpression(true)
          .matchEntireCell(true)
          .findNext();

        if (topSerialCell) {
          currentActiveCount = Number(
            topSerialCell.getDisplayValue() || 0
          );
        }
      }

      studentSerial = String(currentActiveCount + 1);
    }

    var data = {
      '학생일련번호':studentSerial,
      '학생ID':studentId,
      '학생이름':p.studentName || p['학생이름'] || '',
      '비밀번호':password,
      '학습시작일':wmStudentLearningDateValue_(learningStartDate),
      '학습종료일':learningEndDate ? wmStudentLearningDateValue_(learningEndDate) : '',
      '출석약속':String(p.attendancePromise || p['출석약속'] || '').replace(/[^1-7]/g, '').slice(0, 1),
      '세트약속':String(p.setPromise || p['세트약속'] || '').replace(/[^1-7]/g, '').slice(0, 1),
      '최초배정':normalizeLearningAssignForLms_(p.learningAssign || p['최초배정'] || ''),
      '학습배정':normalizeLearningAssignForLms_(p.learningAssign || p['최초배정'] || ''),
      'S4실루엣단계':normalizeS4SilhouetteStep_(p.s4 || p['S4실루엣단계'] || ''),
      'S6테스트모드':normalizeS6TestMode_(p.s6 || p['S6테스트모드'] || ''),
      '레벨완료조건':normalizeLevelCompleteCondition_(p.levelCompleteCondition || p['레벨완료조건'] || ''),
      '현재세트':p.currentSet || p['현재세트'] || '',
      '반명':p.className || p['반명'] || '', 'Class':p.className || p['반명'] || '',
      '교사명':p.teacher || p['교사명'] || '', '학교':p.school || p['학교'] || '',
      '학년':p.grade || p['학년'] || '', '주소':p.address || p['주소'] || '',
      '수신자구분':p.recipientType || p['수신자구분'] || p['수신구분'] || p['수신'] || '',
      '학부모명':p.parentName || p['학부모명'] || '', '학부모연락처':p.parentPhone || p['학부모연락처'] || '',
      '최종결제':p.lastPayment || p['최종결제'] || '', '등록일':registerDateRaw,
      '최초등록일':p.firstRegisterDate || p['최초등록일'] || '', '최근결제일':p.recentPaymentDate || p['최근결제일'] || '',
      '결제공백일수':p.paymentGapDays || p['결제공백일수'] || '', '재등록횟수':p.reregisterCount || p['재등록횟수'] || '',
      '총등록기간':p.totalRegisterPeriod || p['총등록기간'] || '', '등록상태':p.registerStatus || p['등록상태'] || '재원중',
      '활성상태':newActiveStatus, '반ID':p.classId || p['반ID'] || '',
      '교사ID':p.teacherId || p['교사ID'] || '', '현재세션':p.currentSession || p['현재세션'] || '',
      '세션상태':p.sessionStatus || p['세션상태'] || 'LOGOUT', '최종로그인':p.lastLogin || p['최종로그인'] || '',
      '접속기기':p.device || p['접속기기'] || '', '비고':p.memo || p['비고'] || ''
    };
    if (String(data['반명'] || '').trim()) {
      var classValidation = wmValidateStudentClassMoveForLms_(getLmsSpreadsheet_(), data['반명']);
      if (!classValidation.success) return {success:false, message:classValidation.message || '반관리_DB 대상 반 확인 실패'};
    }

    var row = headers.map(function(header){ return Object.prototype.hasOwnProperty.call(data, header) ? data[header] : ''; });
    var savedRow = 2;
    sheet.insertRowAfter(1);
    var parentPhoneColumn = headers.indexOf('학부모연락처');
    if (parentPhoneColumn >= 0) sheet.getRange(savedRow, parentPhoneColumn + 1).setNumberFormat('@');
    var newStudentTeacherIdColumn = findFirstHeaderIndex_(headers, ['교사ID','담당교사ID']);
    if (newStudentTeacherIdColumn >= 0 && String(data['교사ID'] || '').trim()) {
      var newStudentTeacherIdA1 = sheet.getRange(savedRow, newStudentTeacherIdColumn + 1).getA1Notation();
      var newStudentTeacherIdValidation = wmPrepareTeacherIdValidationForLms_(getLmsSpreadsheet_(), sheet, [newStudentTeacherIdA1]);
      if (!newStudentTeacherIdValidation.success) {
        try { sheet.deleteRow(savedRow); } catch (ignoreValidationRollback) {}
        return {success:false, message:newStudentTeacherIdValidation.message || '신규학생 교사ID 저장규칙 확인 실패'};
      }
    }
    sheet.getRange(savedRow, 1, 1, lastCol).setValues([row]);
    var learningStartColumn = headers.indexOf('학습시작일');
    var learningEndColumn = headers.indexOf('학습종료일');
    if (learningStartColumn >= 0) sheet.getRange(savedRow, learningStartColumn + 1).setNumberFormat('yyyy-MM-dd');
    if (learningEndColumn >= 0) sheet.getRange(savedRow, learningEndColumn + 1).setNumberFormat('yyyy-MM-dd');

    var classSync = wmSyncRegisteredStudentToClassDb_(
      getLmsSpreadsheet_(),
      data['반명'],
      studentId,
      data['학생이름'],
      data['교사ID'],
      data['교사명']
    );
    if (!classSync.success) {
      try {
        var savedId = String(sheet.getRange(savedRow, headers.indexOf('학생ID') + 1).getDisplayValue() || '').trim();
        if (savedId.toUpperCase() === studentId.toUpperCase()) sheet.deleteRow(savedRow);
      } catch (rollbackErr) {}
      return {
        success:false,
        message:'반관리_DB 갱신 실패로 학생등록을 취소했습니다.',
        error:classSync.message || '',
        studentId:studentId,
        studentSerial:studentSerial
      };
    }

    wmClearRuntimeCachesForStudent_(studentId);
    var sendCenterSync = {success:true, skipped:true};
    try {
      wmSyncSendCenterPrimaryDataFromStudents_([data]);
      sendCenterSync = {success:true};
    } catch (sendErr) {
      sendCenterSync = {success:false, message:String(sendErr && sendErr.message ? sendErr.message : sendErr)};
    }
    return {
      success:true,
      message:'학생관리_DB 저장완료',
      studentId:studentId,
      studentSerial:studentSerial,
      student:data,
      classSync:classSync,
      sendCenterSync:sendCenterSync
    };
  } catch (err) {
    return {success:false, message:'학생관리_DB 저장 오류', error:String(err && err.message ? err.message : err)};
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getStudentsForLmsApi_(e) {
  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('1.학생관리_DB');
    if (!sheet) return outputResult(e, { success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.', students:[] });

    if (wmIsLmsInitialPageRequest_(e)) {
      var initialStudentResult = wmReadLmsFirstRowsAsObjects_('1.학생관리_DB', wmLmsInitialLimit_(e, 25));
      var initialActor = wmGetLmsRequestActor_(e);
      var initialStudents = wmFilterStudentsForActor_(initialStudentResult.rows || [], initialActor);
      initialStudents = wmRedactStudentSensitiveFieldsForActor_(initialStudents, initialActor);
      return outputResult(e, {success:true, students:initialStudents, total:initialStudentResult.total || initialStudents.length, initialPage:true});
    }

    var values = sheet.getDataRange().getDisplayValues();
    if (!values || !values.length) return outputResult(e, { success:true, students:[] });

    var headers = values[0].map(function(h) { return String(h || '').trim(); });
    var idxOfficialCurrentSet = headers.indexOf('현재세트');
    var idxInitialAssign = findFirstHeaderIndex_(headers, ['최초배정', '학습배정']);
    if (headers.indexOf('최초배정') < 0 && idxInitialAssign >= 0) {
      sheet.getRange(1, idxInitialAssign + 1).setValue('최초배정');
      headers[idxInitialAssign] = '최초배정';
    }
    var currentProgressSetMap = buildCurrentProgressSetMapForLms_();
    var officialCurrentSetValues = idxOfficialCurrentSet >= 0
      ? values.slice(1).map(function(row) { return [String(row[idxOfficialCurrentSet] || '').trim()]; })
      : [];
    var officialCurrentSetChanged = false;
    var students = [];
    for (var r = 1; r < values.length; r++) {
      var rowObj = {};
      for (var c = 0; c < headers.length; c++) {
        if (headers[c]) rowObj[headers[c]] = values[r][c];
      }
      if (idxOfficialCurrentSet >= 0) rowObj['현재세트'] = values[r][idxOfficialCurrentSet];
      if (String(rowObj['학생ID'] || '').trim()) {
        rowObj['학습배정'] = rowObj['최초배정'] || rowObj['학습배정'] || '';
        rowObj['S4실루엣단계'] = normalizeS4SilhouetteStep_(rowObj['S4실루엣단계']);
        rowObj['S6테스트모드'] = normalizeS6TestMode_(rowObj['S6테스트모드'] || rowObj['학습모드']);
        rowObj['레벨완료조건'] = normalizeLevelCompleteCondition_(rowObj['레벨완료조건']);
        rowObj['레벨회차추가'] = normalizeLevelExtraRounds_(rowObj['레벨회차추가']);
        /* WM_LMS_CURRENT_SET_FROM_PROGRESS_DB_20260618_V1
         * LMS 현재세트는 1.학생관리_DB의 현재세트가 아니라
         * 8.현재진행_DB의 학생ID별 최신 Set_ID를 기준으로 표시합니다.
         * 현재진행_DB에 기록이 없으면 빈칸으로 둡니다.
         */
        var lmsStudentId = String(rowObj['학생ID'] || '').trim().toUpperCase();
        var officialCurrentSet = String(currentProgressSetMap[lmsStudentId] || '').trim();
        rowObj['현재세트'] = officialCurrentSet;
        if (idxOfficialCurrentSet >= 0 && officialCurrentSetValues[r - 1][0] !== officialCurrentSet) {
          officialCurrentSetValues[r - 1][0] = officialCurrentSet;
          officialCurrentSetChanged = true;
        }
        students.push(rowObj);
      }
    }

    if (idxOfficialCurrentSet >= 0 && officialCurrentSetChanged && officialCurrentSetValues.length) {
      var officialCurrentSetRange = sheet.getRange(2, idxOfficialCurrentSet + 1, officialCurrentSetValues.length, 1);
      officialCurrentSetRange.setNumberFormat('@');
      officialCurrentSetRange.setValues(officialCurrentSetValues);
    }

    wmSyncSendCenterPrimaryDataFromStudents_(students);

    var actor = wmGetLmsRequestActor_(e);
    students = wmFilterStudentsForActor_(students, actor);

    /* WM_PROFILE_STUDENT_DB_ID_LOOKUP_V1
     * 프로필센터에서는 학생ID를 최종 조회키로 사용합니다.
     * 기존 학생관리 화면의 전체조회는 파라미터가 없으므로 그대로 유지됩니다.
     */
    var requestedStudentId = String(
      e && e.parameter ? (e.parameter.studentId || e.parameter.Student_ID || '') : ''
    ).trim().toUpperCase();
    var requestedStudentName = String(
      e && e.parameter ? (e.parameter.studentName || e.parameter.Student_Name || '') : ''
    ).trim();

    var requestedMatchMode = String(
      e && e.parameter ? (e.parameter.matchMode || 'exact') : 'exact'
    ).trim().toLowerCase();

    if (requestedStudentId || requestedStudentName) {
      students = students.filter(function(student) {
        var rowId = String(student['학생ID'] || '').trim().toUpperCase();
        var rowName = String(student['학생이름'] || '').trim();
        var idMatched = !requestedStudentId || (
          requestedMatchMode === 'contains'
            ? rowId.indexOf(requestedStudentId) !== -1
            : rowId === requestedStudentId
        );
        var nameMatched = !requestedStudentName || (
          requestedMatchMode === 'contains'
            ? rowName.indexOf(requestedStudentName) !== -1
            : rowName === requestedStudentName
        );
        return idMatched && nameMatched;
      });
    }

    students = wmRedactStudentSensitiveFieldsForActor_(students, actor);

    return outputResult(e, { success:true, students:students });
  } catch (err) {
    return outputResult(e, { success:false, message:'학생관리_DB 조회 오류', error:String(err && err.message ? err.message : err), students:[] });
  }
}


/* WM_LMS_FIRST_SCREEN_MINIMUM_ROWS_V1
 * 각 관리 탭 최초 진입은 전체 DB를 읽지 않고 화면에 필요한 선두 행만 읽습니다.
 * 검색/상세조회는 기존 전체조회 경로를 그대로 사용합니다. */
function wmIsLmsInitialPageRequest_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  return String(p.initial || p.firstPage || '').trim() === '1';
}

function wmLmsInitialLimit_(e, fallback) {
  var p = (e && e.parameter) ? e.parameter : {};
  var limit = Number(p.limit || fallback || 25);
  if (!isFinite(limit) || limit < 1) limit = Number(fallback || 25);
  return Math.min(Math.max(Math.floor(limit), 1), 50);
}

function wmReadLmsFirstRowsAsObjects_(sheetName, limit) {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success:false, message:sheetName + ' 시트를 찾을 수 없습니다.', headers:[], rows:[], total:0 };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { success:true, headers:[], rows:[], total:0 };
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
  var count = Math.min(Math.max(lastRow - 1, 0), Number(limit || 25));
  var values = count ? sheet.getRange(2, 1, count, lastCol).getDisplayValues() : [];
  var rows = values.map(function(row){
    var obj = {};
    headers.forEach(function(h, i){ if (h) obj[h] = row[i]; });
    return obj;
  }).filter(function(obj){ return Object.keys(obj).some(function(k){ return String(obj[k] || '').trim(); }); });
  return { success:true, headers:headers, rows:rows, total:Math.max(lastRow - 1, 0), initialPage:true };
}

function wmAllowedStudentIdSetForActor_(actor) {
  if (!actor || !/^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) return null;
  var key = wmCacheKey_('LMS_ALLOWED_STUDENTS', String(actor.teacherId || actor.name || ''));
  var cached = wmCacheGetJson_(key);
  if (cached && cached.ids) return cached.ids;
  var students = wmReadLmsSheetAsObjects_('1.학생관리_DB').rows || [];
  var ids = wmStudentIdSet_(wmFilterStudentsForActor_(students, actor));
  wmCachePutJson_(key, {ids:ids}, 120);
  return ids;
}

function wmReadLmsSheetAsObjects_(sheetName) {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success:false, message:sheetName + ' 시트를 찾을 수 없습니다.', headers:[], rows:[] };
  var values = sheet.getDataRange().getDisplayValues();
  if (!values || !values.length) return { success:true, headers:[], rows:[] };
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    var hasValue = false;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = values[r][c];
      if (String(values[r][c] || '').trim()) hasValue = true;
    }
    if (hasValue) rows.push(obj);
  }
  if (sheetName === '1.학생관리_DB') {
    wmSyncSendCenterPrimaryDataFromStudents_(rows, ss);
  }
  return { success:true, headers:headers, rows:rows };
}


/* WM_GRADE_LEVEL_DB_LMS_API_V1
 * 학생관리/학습기록과 동일한 LMS 조회 구조를 사용합니다.
 * 원본 시트 6.성적등급_DB의 헤더명과 표시값을 그대로 객체로 반환합니다.
 */
function wmGradeGuideText_() {
  return [
    '[성적등급산정법]',
    '출석점수·학업집중도·점수추이 평가는',
    '각각 10점 만점이며, 성적등급은',
    '이를 합산한 통합점수에 따라 1.1등급(최상)부터',
    '3.10등급(최하)까지 구분됩니다.'
  ].join('\n');
}

function wmGetGradeLevelsForLmsApi_(e) {
  try {
    var result = wmIsLmsInitialPageRequest_(e)
      ? wmReadLmsFirstRowsAsObjects_('6.성적등급_DB', wmLmsInitialLimit_(e, 25))
      : wmReadLmsSheetAsObjects_('6.성적등급_DB');
    var rows = result.rows || [];
    var actor = wmGetLmsRequestActor_(e);

    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      rows = wmFilterRowsByStudentSet_(rows, wmAllowedStudentIdSetForActor_(actor) || {}, actor);
    }

    return outputResult(e, {
      success: !!result.success,
      message: result.message || '',
      gradeLevels: rows,
      gradeLevelRows: rows,
      rows: rows,
      gradeGuide: wmGradeGuideText_()
    });
  } catch (err) {
    return outputResult(e, {
      success:false,
      message:'6.성적등급_DB 조회 오류',
      error:String(err && err.message ? err.message : err),
      gradeLevels:[],
      gradeLevelRows:[],
      rows:[]
    });
  }
}


/* WM_REPORT_INITIAL_IMPORT_V1 */
function wmGetReportsForLmsApi_(e) {
  try {
    var result = wmIsLmsInitialPageRequest_(e) ? wmReadLmsFirstRowsAsObjects_('4-1.성적표생성_DB', wmLmsInitialLimit_(e, 25)) : wmReadLmsSheetAsObjects_('4-1.성적표생성_DB');
    var rows = result.rows || [];
    var actor = wmGetLmsRequestActor_(e);
    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      rows = wmFilterRowsByStudentSet_(rows, wmAllowedStudentIdSetForActor_(actor) || {}, actor);
    }

    /* 발행주차로 유형과 발행시점을 확정합니다.
     * 마지막 주에 실제 학습기록이 없어도 마지막 주차 행은 Monthly로 해석합니다. */
    rows = rows.filter(function(row) {
      var rowInfo = wmReportRowPeriodInfo_(row);
      if (!rowInfo || !wmReportIsReleased_(rowInfo.periodInfo)) return false;
      row['리포트유형'] = rowInfo.reportType;
      row['비고'] = rowInfo.reportType === '월간' ? 'Monthly Report' : 'Weekly Report';
      return true;
    });

    /* WM_PROFILE_REPORT_DB_LOOKUP_V1
     * 프로필센터 최근 성적표는 복수 학생ID와 기간을 서버에서 먼저 필터링합니다.
     * 기간 미입력 시 오늘 기준 최근 1개월을 기본값으로 사용합니다.
     */
    var requestedStudentIdsText = String(
      e && e.parameter ? (
        e.parameter.studentIds ||
        e.parameter.studentId ||
        e.parameter.Student_ID ||
        ''
      ) : ''
    ).trim();

    var requestedStudentIds = {};
    requestedStudentIdsText.split(',').forEach(function(value) {
      var id = String(value || '').trim().toUpperCase();
      if (id) requestedStudentIds[id] = true;
    });

    function wmProfileReportParseDate_(value) {
      var text = String(value || '').trim();
      var match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!match) return null;
      var date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
      date.setHours(0, 0, 0, 0);
      return isNaN(date.getTime()) ? null : date;
    }

    var requestedStart = wmProfileReportParseDate_(
      e && e.parameter ? e.parameter.dateStart : ''
    );
    var requestedEnd = wmProfileReportParseDate_(
      e && e.parameter ? e.parameter.dateEnd : ''
    );

    if (!requestedStart && !requestedEnd) {
      requestedEnd = new Date();
      requestedEnd.setHours(0, 0, 0, 0);
      requestedStart = new Date(requestedEnd);
      requestedStart.setMonth(requestedStart.getMonth() - 1);
    } else {
      if (!requestedStart) requestedStart = new Date(requestedEnd);
      if (!requestedEnd) {
        requestedEnd = new Date();
        requestedEnd.setHours(0, 0, 0, 0);
      }
    }

    var hasStudentFilter = Object.keys(requestedStudentIds).length > 0;
    rows = rows.filter(function(row) {
      var rowId = String(row['학생ID'] || '').trim().toUpperCase();
      if (hasStudentFilter && !requestedStudentIds[rowId]) return false;

      var filterRowInfo = wmReportRowPeriodInfo_(row);
      var rowDate = wmProfileReportParseDate_(
        filterRowInfo && filterRowInfo.periodInfo ? filterRowInfo.periodInfo.endDate : ''
      );
      if (!rowDate) return false;
      if (requestedStart && rowDate < requestedStart) return false;
      if (requestedEnd && rowDate > requestedEnd) return false;
      return true;
    });

    return outputResult(e, {
      success:!!result.success,
      message:result.message || '',
      reports:rows,
      reportRows:rows,
      rows:rows
    });
  } catch (err) {
    return outputResult(e, {
      success:false,
      message:'4-1.성적표생성_DB 조회 오류',
      error:String(err && err.message ? err.message : err),
      reports:[],
      reportRows:[],
      rows:[]
    });
  }
}

function wmReportCell_(row, names) {
  row = row || {}; names = names || [];
  for (var i = 0; i < names.length; i++) {
    var key = String(names[i] || '').trim();
    if (key && row.hasOwnProperty(key) && String(row[key] || '').trim() !== '') return row[key];
  }
  return '';
}

function wmReportNormalizeDate_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  var date = new Date(text);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
}

function wmReportWeekKey_(dateText) {
  var match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return '';
  return match[1] + '-' + match[2] + '-' + wmReportCalendarWeek_(dateText) + '주차';
}

/* 월요일~일요일 달력 주차를 사용합니다.
 * 월 경계 기록은 학습일의 월에만 포함하고, 6주차가 생기면 5주차에 병합합니다.
 */
function wmReportCalendarWeek_(dateText) {
  var match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 0;
  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);
  var firstDay = new Date(year, month - 1, 1);
  var mondayOffset = (firstDay.getDay() + 6) % 7;
  var calendarWeek = Math.floor((mondayOffset + day - 1) / 7) + 1;
  return Math.max(1, Math.min(5, calendarWeek));
}

function wmReportCalendarWeekRange_(yearMonth, week) {
  var match = String(yearMonth || '').match(/^(\d{4})-(\d{2})$/);
  if (!match) return {startDate:'',endDate:'',lastWeek:0,isLastWeek:false};
  var year = Number(match[1]);
  var month = Number(match[2]);
  var lastDay = new Date(year, month, 0).getDate();
  var firstDay = new Date(year, month - 1, 1);
  var mondayOffset = (firstDay.getDay() + 6) % 7;
  var lastIso = match[1] + '-' + match[2] + '-' + ('0' + lastDay).slice(-2);
  var lastWeek = wmReportCalendarWeek_(lastIso);
  var normalizedWeek = Math.max(1, Math.min(lastWeek, Number(week || 1)));
  var startDay = normalizedWeek === 1 ? 1 : 1 + (7 - mondayOffset) + (normalizedWeek - 2) * 7;
  var endDay = normalizedWeek === 1
    ? Math.min(lastDay, 7 - mondayOffset)
    : (normalizedWeek === lastWeek ? lastDay : Math.min(lastDay, startDay + 6));
  if (startDay > lastDay) return {startDate:'',endDate:'',lastWeek:lastWeek,isLastWeek:false};
  var prefix = match[1] + '-' + match[2] + '-';
  return {
    startDate:prefix + ('0' + startDay).slice(-2),
    endDate:prefix + ('0' + endDay).slice(-2),
    lastWeek:lastWeek,
    isLastWeek:normalizedWeek === lastWeek
  };
}

/* WM_REPORT_RELEASE_POLICY_V2
 * Weekly: 해당 월 1일부터 선택한 종료 주차 말일까지 2.학습기록_DB 누적기록 중 완료기록만 집계하며, 마지막 주 Weekly는 미발행합니다.
 * Monthly: 다음 달 1일 00:01 이후 직전 월 1일부터 말일까지 2.학습기록_DB 누적기록 중 완료기록만 집계합니다.
 * 주차는 최대 5주차이며 달력상 6주차는 5주차에 병합합니다.
 */
function wmReportSeoulDateKey_(date) {
  return Utilities.formatDate(date || new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function wmReportPeriodInfo_(dateText) {
  var match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  var yearMonth = match[1] + '-' + match[2];
  var week = wmReportCalendarWeek_(dateText);
  var range = wmReportCalendarWeekRange_(yearMonth, week);
  return {
    year:Number(match[1]), month:Number(match[2]), yearMonth:yearMonth, week:week,
    startDate:range.startDate, endDate:range.endDate,
    lastWeek:Number(range.lastWeek || week), isLastWeek:!!range.isLastWeek
  };
}

/* 4-1.성적표생성_DB 행의 발행기간과 유형은 마지막 학습날짜가 아니라
 * 확정된 리포트주차만으로 판정합니다. */
function wmReportRowPeriodInfo_(row) {
  row = row || {};
  var weekText = String(row['리포트주차'] || '').trim();
  var match = weekText.match(/^(\d{4})-(\d{1,2})-(\d+)주차$/);
  if (!match) return null;
  var year = Number(match[1]);
  var month = Number(match[2]);
  var week = Number(match[3]);
  if (!year || month < 1 || month > 12 || week < 1 || week > 5) return null;
  var yearMonth = year + '-' + ('0' + month).slice(-2);
  var monthLastDay = new Date(year, month, 0).getDate();
  var monthEndInfo = wmReportPeriodInfo_(yearMonth + '-' + ('0' + monthLastDay).slice(-2));
  var lastWeek = Number(monthEndInfo && monthEndInfo.lastWeek || 0);
  if (!lastWeek || week > lastWeek) return null;
  var range = wmReportCalendarWeekRange_(yearMonth, week);
  var isLastWeek = week === lastWeek;
  return {
    year:year,
    month:month,
    yearMonth:yearMonth,
    week:week,
    reportType:isLastWeek ? '월간' : '주간',
    periodInfo:{
      year:year,
      month:month,
      yearMonth:yearMonth,
      week:week,
      startDate:range.startDate,
      endDate:range.endDate,
      lastWeek:lastWeek,
      isLastWeek:isLastWeek
    }
  };
}

function wmReportSelectExactRow_(rows, studentId, yearMonth, week, reportType) {
  var expectedStudentId = String(studentId || '').trim().toUpperCase();
  var expectedYearMonth = String(yearMonth || '').trim();
  var expectedWeek = Number(week || 0);
  var expectedType = /월간|monthly/i.test(String(reportType || '')) ? '월간' : '주간';
  for (var i = 0; i < (rows || []).length; i++) {
    var row = rows[i] || {};
    var rowStudentId = String(row['학생ID'] || row['Student_ID'] || row['studentId'] || '').trim().toUpperCase();
    var rowInfo = wmReportRowPeriodInfo_(row);
    if (rowStudentId === expectedStudentId && rowInfo &&
        rowInfo.yearMonth === expectedYearMonth && rowInfo.week === expectedWeek &&
        rowInfo.reportType === expectedType) return row;
  }
  return null;
}

function wmReportIsReleased_(periodInfo, now) {
  if (!periodInfo || !periodInfo.endDate) return false;
  var match = String(periodInfo.endDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;
  /* 종료일 다음 날 00:01(Asia/Seoul)에만 발행 완료로 인정합니다. */
  var releaseAt = Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 15, 1, 0, 0);
  return (now || new Date()).getTime() >= releaseAt;
}

function wmReportParseNumber_(value) {
  var match = String(value || '').replace(/,/g, '').match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function wmReportParseSeconds_(value) {
  var text = String(value || '').trim();
  if (!text) return 0;
  if (/^\d+(?:\.\d+)?$/.test(text)) return Math.max(0, Math.round(Number(text)));
  var h = 0, m = 0, s = 0, match;
  match = text.match(/(\d+(?:\.\d+)?)\s*시간/); if (match) h = Number(match[1]);
  match = text.match(/(\d+(?:\.\d+)?)\s*분/); if (match) m = Number(match[1]);
  match = text.match(/(\d+(?:\.\d+)?)\s*초/); if (match) s = Number(match[1]);
  if (!h && !m && !s && text.indexOf(':') !== -1) {
    var parts = text.split(':').map(function(v){ return Number(v || 0); });
    if (parts.length === 3) return Math.round(parts[0] * 3600 + parts[1] * 60 + parts[2]);
    if (parts.length === 2) return Math.round(parts[0] * 60 + parts[1]);
  }
  return Math.max(0, Math.round(h * 3600 + m * 60 + s));
}

function wmReportFormatDuration_(seconds) {
  var total = Math.max(0, Math.round(Number(seconds || 0)));
  var h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
  if (h > 0) return h + '시간 ' + m + '분' + (s ? ' ' + s + '초' : '');
  if (m > 0) return m + '분' + (s ? ' ' + s + '초' : '');
  return s + '초';
}

function wmReportPlainSetId_(value) {
  return String(value || '').trim().replace(/^WM/i, '');
}

function wmReportStoredSetId_(value) {
  var raw = String(value || '').trim()
    .replace(/[‐‑‒–—―−－]/g, '-')
    .replace(/[./]/g, '-')
    .replace(/\s+/g, '');

  var match = raw.match(/^WM(\d{1,2})-(\d{1,2})-(\d{1,2})$/i);
  if (match) {
    var wmLevel = Number(match[1]);
    if (wmLevel >= 3 && wmLevel <= 13) {
      return 'WM' + wmLevel + '-' + Number(match[2]) + '-' + Number(match[3]);
    }
  }

  match = raw.match(/^20(\d{2})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    var dateLikeLevel = Number(match[1]);
    if (dateLikeLevel >= 3 && dateLikeLevel <= 13) {
      return 'WM' + dateLikeLevel + '-' + Number(match[2]) + '-' + Number(match[3]);
    }
  }

  match = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    var plainLevel = Number(match[1]);
    if (plainLevel >= 3 && plainLevel <= 13) {
      return 'WM' + plainLevel + '-' + Number(match[2]) + '-' + Number(match[3]);
    }
  }

  return raw;
}

function wmReportLevelFromSet_(value) {
  var match = wmReportPlainSetId_(value).match(/^(\d+)-/);
  return match ? Number(match[1]) : 0;
}

function wmReportWordCountForSet_(setId) {
  var level = wmReportLevelFromSet_(setId);
  if (level >= 3 && level <= 6) return 10;
  if (level >= 7 && level <= 13) return 20;
  return 0;
}

function wmReportLevelLabel_(setId, fallback) {
  var level = wmReportLevelFromSet_(setId);
  if (!level) {
    var match = String(fallback || '').match(/(\d+)/);
    level = match ? Number(match[1]) : 0;
  }
  return level ? 'Basic ' + level : String(fallback || '');
}

function wmReportIsCompletedRecord_(row) {
  var status = String(wmReportCell_(row, ['완료상태','완료']) || '').trim();
  if (!status || /미완료|incomplete|fail|중단|학습중/i.test(status)) return false;
  return /완료|complete|completed|success|1회차|2회차|3회차/i.test(status);
}

function wmReportNextId_(existingIds, yearMonth, sequenceState) {
  var prefix = 'R' + String(yearMonth || '').replace('-', '').slice(2);
  if (!sequenceState[prefix]) {
    var max = 0;
    for (var i = 0; i < existingIds.length; i++) {
      var id = String(existingIds[i] || '').trim();
      var match = id.match(new RegExp('^' + prefix + '(\\d{3})$'));
      if (match) max = Math.max(max, Number(match[1]));
    }
    sequenceState[prefix] = max;
  }
  sequenceState[prefix] += 1;
  return prefix + ('000' + sequenceState[prefix]).slice(-3);
}

/* WM_PUBLIC_REPORT_URL_V1
 * 4-1.성적표생성_DB의 기존 '성적표보기' 컬럼을 공개 성적표 URL 저장칸으로 사용합니다.
 * 학생ID·이름·연월은 URL에 노출하지 않고 8자리 랜덤 영문+숫자 토큰만 사용합니다.
 * 기존 성적표 계산·Renderer·Map·LMS 기능은 변경하지 않습니다.
 */
var WM_PUBLIC_REPORT_BASE_URL_ = WM_ENV === 'REAL'
  ? 'https://word-mate-real.vercel.app/report/'
  : 'https://word-mate-test.vercel.app/report/';

function wmNormalizePublicReportToken_(value) {
  var token = String(value || '').trim().toUpperCase();
  return /^[A-Z0-9]{8}$/.test(token) ? token : '';
}

function wmPublicReportTokenFromLink_(value) {
  var text = String(value || '').trim();
  var match = text.match(/\/report\/([A-Z0-9]{8})(?:[\/?#]|$)/i);
  return match ? wmNormalizePublicReportToken_(match[1]) : '';
}

function wmCreatePublicReportToken_(usedTokens) {
  usedTokens = usedTokens || {};
  var alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var attempt = 0; attempt < 50; attempt++) {
    var seed = Utilities.getUuid() + '|' + new Date().getTime() + '|' + Math.random();
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      seed,
      Utilities.Charset.UTF_8
    );
    var token = '';
    for (var i = 0; i < 8; i++) {
      token += alphabet[(digest[i] & 255) % alphabet.length];
    }
    if (!usedTokens[token]) {
      usedTokens[token] = true;
      return token;
    }
  }
  throw new Error('성적표 URL 토큰 생성에 실패했습니다.');
}

function wmNextPublicReportUrl_(usedTokens) {
  return WM_PUBLIC_REPORT_BASE_URL_ + wmCreatePublicReportToken_(usedTokens);
}

function wmPreparePublicReportUrlState_(reportSheet, reportHeaders, reportValues) {
  var state = {success:false, usedTokens:{}, updatedCount:0};
  try {
    var linkIndex = (reportHeaders || []).indexOf('성적표보기');
    var reportIdIndex = (reportHeaders || []).indexOf('리포트ID');
    var studentIdIndex = (reportHeaders || []).indexOf('학생ID');
    var weekIndex = (reportHeaders || []).indexOf('리포트주차');
    if (!reportSheet || linkIndex < 0 || reportIdIndex < 0 || studentIdIndex < 0 || weekIndex < 0 || !reportValues || reportValues.length < 2) {
      return state;
    }

    for (var r = 1; r < reportValues.length; r++) {
      var existingToken = wmPublicReportTokenFromLink_(reportValues[r][linkIndex]);
      if (existingToken) state.usedTokens[existingToken] = true;
    }

    var linkValues = [];
    var changed = false;
    for (var rowIndex = 1; rowIndex < reportValues.length; rowIndex++) {
      var currentLink = String(reportValues[rowIndex][linkIndex] || '').trim();
      var reportId = String(reportValues[rowIndex][reportIdIndex] || '').trim();
      var studentId = String(reportValues[rowIndex][studentIdIndex] || '').trim();
      var weekKey = String(reportValues[rowIndex][weekIndex] || '').trim();
      if (reportId && studentId && weekKey && !wmPublicReportTokenFromLink_(currentLink) && (!currentLink || currentLink.toUpperCase() === 'Y')) {
        reportValues[rowIndex][linkIndex] = wmNextPublicReportUrl_(state.usedTokens);
        state.updatedCount += 1;
        changed = true;
      }
      linkValues.push([reportValues[rowIndex][linkIndex]]);
    }

    if (changed) {
      reportSheet
        .getRange(2, linkIndex + 1, linkValues.length, 1)
        .setNumberFormat('@')
        .setValues(linkValues);
    }

    state.success = true;
    return state;
  } catch (err) {
    state.error = String(err && err.message ? err.message : err);
    return state;
  }
}

function wmGetPublicReportByToken_(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var token = wmNormalizePublicReportToken_(p.token || p.reportToken || '');
    if (!token) return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};

    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('4-1.성적표생성_DB');
    if (!sheet || sheet.getLastRow() < 2) {
      return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
      .map(function(v){ return String(v || '').trim(); });
    var linkIndex = headers.indexOf('성적표보기');
    var studentIndex = headers.indexOf('학생ID');
    var weekIndex = headers.indexOf('리포트주차');
    var reportIdIndex = headers.indexOf('리포트ID');
    if (linkIndex < 0 || studentIndex < 0 || weekIndex < 0) {
      return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};
    }

    var rowCount = sheet.getLastRow() - 1;
    var links = sheet.getRange(2, linkIndex + 1, rowCount, 1).getDisplayValues();
    var targetRow = 0;
    for (var i = 0; i < links.length; i++) {
      if (wmPublicReportTokenFromLink_(links[i][0]) === token) {
        targetRow = i + 2;
        break;
      }
    }
    if (!targetRow) return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};

    var values = sheet.getRange(targetRow, 1, 1, lastCol).getDisplayValues()[0];
    var studentId = String(values[studentIndex] || '').trim().toUpperCase();
    var weekKey = String(values[weekIndex] || '').trim();
    var weekMatch = weekKey.match(/^(\d{4}-\d{2})-(\d+)주차$/);
    if (!studentId || !weekMatch) {
      return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};
    }

    var reportResult = wmGetMonthlyReportDataForLms_({
      parameter:{
        studentId:studentId,
        yearMonth:weekMatch[1],
        reportWeek:weekKey,
        requestSource:'publicReport'
      }
    });
    if (!reportResult || reportResult.success !== true || !reportResult.data) {
      return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};
    }

    return {
      success:true,
      reportId:reportIdIndex >= 0 ? String(values[reportIdIndex] || '').trim() : '',
      data:reportResult.data
    };
  } catch (err) {
    return {success:false, message:'성적표를 확인할 수 없습니다.', data:null};
  }
}

function wmInitializeReportsFromLearningRecordsForLms_(e) {
  try {
    var ss = getLmsSpreadsheet_();
    var reportSheet = ss.getSheetByName('4-1.성적표생성_DB');
    if (!reportSheet) return {success:false, message:'4-1.성적표생성_DB 시트를 찾을 수 없습니다.'};

    var requiredHeaders = ['리포트ID','리포트주차','학생ID','학교','학년','학생이름','학습레벨','Set_ID','단어수','학습날짜','점수','총소요시간','성적표보기','Teacher’s Comment','격려메시지','로고URL','비고'];
    var reportValues = reportSheet.getDataRange().getDisplayValues();
    var reportHeaders = reportValues.length ? reportValues[0].map(function(h){ return String(h || '').trim(); }) : [];
    var missing = requiredHeaders.filter(function(h){ return reportHeaders.indexOf(h) < 0; });
    if (missing.length) return {success:false, message:'성적표생성_DB 필수 컬럼 누락: ' + missing.join(', ')};
    var publicUrlState = wmPreparePublicReportUrlState_(reportSheet, reportHeaders, reportValues);

    var recordsResult = wmReadLmsSheetAsObjects_('2.학습기록_DB');
    var studentsResult = wmReadLmsSheetAsObjects_('1.학생관리_DB');
    if (!recordsResult.success) return {success:false, message:recordsResult.message || '학습기록_DB 조회 실패'};

    var actor = wmGetLmsRequestActor_(e);
    var actorRole = wmNormalizeLmsRole_(actor && actor.role);
    var actorRestricted = !!(actor && actor.found) &&
      /^(TEACHER|SUB_LEADER)$/.test(actorRole);
    if (actorRestricted && !studentsResult.success) {
      return {success:false, message:studentsResult.message || '학생관리_DB 조회 실패'};
    }

    var students = studentsResult.success ? (studentsResult.rows || []) : [];
    if (actorRestricted) students = wmFilterStudentsForActor_(students, actor);

    var studentMap = {}, allowedIds = {};
    students.forEach(function(student){
      var sid = String(wmReportCell_(student, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase();
      if (!sid) return;
      studentMap[sid] = student;
      allowedIds[sid] = true;
    });

    var requestedStudentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID)) || '').trim().toUpperCase();
    if (actorRestricted && requestedStudentId && !allowedIds[requestedStudentId]) {
      return {success:false, message:'학생관리_DB에서 대상 학생을 찾을 수 없거나 조회 권한이 없습니다: ' + requestedStudentId};
    }

    var existingKeys = {}, existingIds = [];
    for (var r = 1; r < reportValues.length; r++) {
      var obj = {};
      for (var c = 0; c < reportHeaders.length; c++) obj[reportHeaders[c]] = reportValues[r][c];
      var existingSid = String(obj['학생ID'] || '').trim().toUpperCase();
      var existingWeek = String(obj['리포트주차'] || '').trim();
      if (existingSid && existingWeek) existingKeys[existingSid + '|' + existingWeek] = true;
      if (String(obj['리포트ID'] || '').trim()) existingIds.push(String(obj['리포트ID']).trim());
    }

    var groups = {};
    var wmReportRecoveryDiag = {
      recordCount:(recordsResult.rows || []).length,
      studentCount:students.length,
      allowedStudentCount:Object.keys(allowedIds).length,
      existingReportKeyCount:Object.keys(existingKeys).length,
      missingStudentId:0,
      studentNotAllowed:0,
      requestedStudentMismatch:0,
      invalidDate:0,
      invalidPeriod:0,
      invalidWeekKey:0,
      missingSetId:0,
      notReleased:0,
      acceptedRecords:0
    };

    function addRecordToReportGroup_(key, sid, weekKey, reportType, dateText, setId, record) {
      if (!groups[key]) groups[key] = {sid:sid, weekKey:weekKey, reportType:reportType, latestDate:'', latestSetId:'', latestRecord:null, scoreSum:0, scoreCount:0, totalSeconds:0, wordCount:0};
      var group = groups[key];
      if (!group.latestDate || dateText >= group.latestDate) {
        group.latestDate = dateText;
        group.latestSetId = setId;
        group.latestRecord = record;
      }
      var scoreText = wmReportCell_(record, ['점수','한영주관식']);
      var score = wmReportParseNumber_(scoreText);
      if (String(scoreText || '').trim() !== '' && isFinite(score)) {
        group.scoreSum += score;
        group.scoreCount += 1;
      }
      group.totalSeconds += wmReportParseSeconds_(wmReportCell_(record, ['총소요시간','총 소요시간']));
      group.wordCount += wmReportWordCountForSet_(setId);
    }

    (recordsResult.rows || []).forEach(function(record){
      var sid = String(wmReportCell_(record, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase();
      if (!sid) { wmReportRecoveryDiag.missingStudentId += 1; return; }
      if (actorRestricted && !allowedIds[sid]) { wmReportRecoveryDiag.studentNotAllowed += 1; return; }
      if (requestedStudentId && sid !== requestedStudentId) { wmReportRecoveryDiag.requestedStudentMismatch += 1; return; }
      if (!wmReportIsCompletedRecord_(record)) return;

      var dateText = wmReportNormalizeDate_(wmReportCell_(record, ['학습날짜','학습일','날짜','등록일']));
      if (!dateText) { wmReportRecoveryDiag.invalidDate += 1; return; }
      var reportStudent = studentMap[sid];
      if (!reportStudent) return;
      var reportLearningStart = wmReportNormalizeDate_(wmReportCell_(reportStudent, ['학습시작일'])) ||
        wmReportNormalizeDate_(wmReportCell_(reportStudent, ['등록일']));
      var reportLearningEnd = wmReportNormalizeDate_(wmReportCell_(reportStudent, ['학습종료일']));
      if ((reportLearningStart && dateText < reportLearningStart) ||
          (reportLearningEnd && dateText > reportLearningEnd)) return;
      var periodInfo = wmReportPeriodInfo_(dateText);
      if (!periodInfo) { wmReportRecoveryDiag.invalidPeriod += 1; return; }
      var weekKey = wmReportWeekKey_(dateText);
      if (!weekKey) { wmReportRecoveryDiag.invalidWeekKey += 1; return; }
      var setId = wmReportPlainSetId_(wmReportCell_(record, ['Set_ID','Set ID','세트ID']));
      if (!setId) { wmReportRecoveryDiag.missingSetId += 1; return; }

      var accepted = false;
      var yearMonth = dateText.slice(0, 7);
      var yearMonthParts = yearMonth.split('-');
      var monthLastDay = new Date(Number(yearMonthParts[0]), Number(yearMonthParts[1]), 0).getDate();
      var monthEndInfo = wmReportPeriodInfo_(yearMonth + '-' + ('0' + monthLastDay).slice(-2));

      /* Weekly는 2.학습기록_DB의 누적기록에서 해당 월 1일부터
         종료된 선택 주차까지의 완료기록을 같은 주차 행에 반영합니다. */
      var lastWeek = Number(monthEndInfo && monthEndInfo.lastWeek || 0);
      for (var targetWeek = Number(periodInfo.week || 1); targetWeek < lastWeek; targetWeek++) {
        var targetRange = wmReportCalendarWeekRange_(yearMonth, targetWeek);
        var targetPeriodInfo = wmReportPeriodInfo_(targetRange.endDate);
        if (!targetPeriodInfo || targetPeriodInfo.isLastWeek || !wmReportIsReleased_(targetPeriodInfo)) continue;
        var targetWeekKey = yearMonth + '-' + targetWeek + '주차';
        addRecordToReportGroup_(sid + '|' + targetWeekKey, sid, targetWeekKey, '주간', dateText, setId, record);
        accepted = true;
      }

      /* 마지막 주는 별도 Weekly를 만들지 않습니다.
         월 종료 후 다음 달 1일 00:01부터 해당 월의 누적기록 전체를 마지막 주차 Monthly 한 행에 반영합니다. */
      if (monthEndInfo && wmReportIsReleased_(monthEndInfo)) {
        var monthlyWeekKey = yearMonth + '-' + Number(monthEndInfo.lastWeek || monthEndInfo.week || 1) + '주차';
        addRecordToReportGroup_(sid + '|' + monthlyWeekKey, sid, monthlyWeekKey, '월간', dateText, setId, record);
        accepted = true;
      }

      if (!accepted) { wmReportRecoveryDiag.notReleased += 1; return; }
      wmReportRecoveryDiag.acceptedRecords += 1;
    });

    var rowsToAppend = [], createdCount = 0, skippedCount = 0, studentSet = {}, sequenceState = {};
    Object.keys(groups).sort().forEach(function(key){
      var group = groups[key];
      studentSet[group.sid] = true;
      if (existingKeys[key]) { skippedCount += 1; return; }

      var student = studentMap[group.sid] || {};
      var sourceRecord = group.latestRecord || {};
      var reportId = wmReportNextId_(existingIds, group.latestDate.slice(0,7), sequenceState);
      existingIds.push(reportId);

      var rowObj = {
        '리포트ID':reportId,
        '리포트주차':group.weekKey,
        '학생ID':group.sid,
        '학교':wmReportCell_(student, ['학교','학교명']) || wmReportCell_(sourceRecord, ['학교','학교명']),
        '학년':wmReportCell_(student, ['학년']) || wmReportCell_(sourceRecord, ['학년']),
        '학생이름':wmReportCell_(student, ['학생이름','학생명']) || wmReportCell_(sourceRecord, ['학생이름','학생명']),
        '학습레벨':wmReportLevelLabel_(group.latestSetId, wmReportCell_(student, ['최초배정','학습배정']) || wmReportCell_(sourceRecord, ['학습레벨','현재레벨','레벨'])),
        'Set_ID':wmReportStoredSetId_(group.latestSetId),
        '단어수':group.wordCount,
        '학습날짜':group.latestDate,
        '점수':group.scoreCount ? Math.round(group.scoreSum / group.scoreCount) + '점' : '',
        '총소요시간':wmReportFormatDuration_(group.totalSeconds),
        '성적표보기': publicUrlState.success ? wmNextPublicReportUrl_(publicUrlState.usedTokens) : 'Y',
        'Teacher’s Comment':'',
        '격려메시지':'',
        '로고URL':'',
        '비고':'기존 학습기록 초기생성'
      };

      rowsToAppend.push(reportHeaders.map(function(header){ return rowObj.hasOwnProperty(header) ? rowObj[header] : ''; }));
      existingKeys[key] = true;
      createdCount += 1;
    });

    if (rowsToAppend.length) {
      var startRow = reportSheet.getLastRow() + 1;
      reportSheet.getRange(startRow, reportHeaders.indexOf('리포트ID') + 1, rowsToAppend.length, 1).setNumberFormat('@');
      reportSheet.getRange(startRow, reportHeaders.indexOf('학생ID') + 1, rowsToAppend.length, 1).setNumberFormat('@');
      reportSheet.getRange(startRow, reportHeaders.indexOf('Set_ID') + 1, rowsToAppend.length, 1).setNumberFormat('@');
      reportSheet.getRange(startRow, 1, rowsToAppend.length, reportHeaders.length).setValues(rowsToAppend);
    }

    wmReportRecoveryDiag.groupCount = Object.keys(groups).length;
    wmReportRecoveryDiag.createdCount = createdCount;
    wmReportRecoveryDiag.skippedCount = skippedCount;
    wmReportRecoveryDiag.groupKeys = Object.keys(groups).slice(0, 30);
    Logger.log('WM_REPORT_RECOVERY_DIAG=' + JSON.stringify(wmReportRecoveryDiag));
    return {success:true, message:'기존 학습기록 초기생성 완료', studentCount:Object.keys(studentSet).length, groupCount:Object.keys(groups).length, createdCount:createdCount, skippedCount:skippedCount, requestedStudentId:requestedStudentId, diagnostic:wmReportRecoveryDiag};
  } catch (err) {
    return {success:false, message:'기존 학습기록 성적표 초기생성 오류', error:String(err && err.message ? err.message : err)};
  }
}

/* WM_MONTHLY_REPORT_DATA_API_V1
 * 177개 성적표 바인딩용 통합 원천.
 * 학생ID + 연월 기준으로 학생/학습기록/성적표/현재진행 DB를 조회·계산합니다.
 */
function wmMonthlyReportParam_(e, names) {
  var p = e && e.parameter ? e.parameter : {};
  for (var i = 0; i < names.length; i++) {
    var value = String(p[names[i]] || '').trim();
    if (value) return value;
  }
  return '';
}

/* WM_LMS_RECENT_REPORT_MONTHS_V2
 * 최근성적표보기는 실제 발행이 끝난 성적표 기간만 4개월 목록으로 제공합니다.
 * 현재 월은 종료·발행된 일반 Weekly가 있을 때만 포함하며, 진행 중인 이번 주는 절대 포함하지 않습니다.
 * 월 마지막 주는 별도 Weekly를 발행하지 않고 다음 달 1일 00:01 Monthly에 포함합니다.
 * 달력상 6주차 기록은 기존 기준대로 5주차 Monthly에 통합합니다.
 */
function wmGetRecentReportMonthsForLms_() {
  var timeZone = Session.getScriptTimeZone() || 'Asia/Seoul';
  var now = new Date();
  var todayText = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
  var nowTime = Utilities.formatDate(now, timeZone, 'HH:mm');
  var parts = todayText.split('-');
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var day = Number(parts[2]);
  var currentYearMonth = parts[0] + '-' + parts[1];
  var currentMonthEnd = currentYearMonth + '-' + ('0' + new Date(year, month, 0).getDate()).slice(-2);
  var currentMonthEndInfo = wmReportPeriodInfo_(currentMonthEnd);
  var currentLastWeek = Math.max(1, Math.min(5, Number(currentMonthEndInfo && currentMonthEndInfo.lastWeek || 1)));
  var latestReleasedWeek = 0;

  for (var weekNo = 1; weekNo < currentLastWeek; weekNo++) {
    var weekRange = wmReportCalendarWeekRange_(currentYearMonth, weekNo);
    var weekPeriod = wmReportPeriodInfo_(weekRange.endDate);
    if (weekPeriod && wmReportIsReleased_(weekPeriod, now)) latestReleasedWeek = weekNo;
  }

  /* 현재 월에 발행 완료 Weekly가 없으면 최근 성적표의 기준 월은 직전 월입니다. */
  var anchorOffset = latestReleasedWeek > 0 ? 0 : 1;
  var months = [];

  for (var listIndex = 0; listIndex < 4; listIndex++) {
    var offset = anchorOffset + listIndex;
    var absoluteMonth = year * 12 + (month - 1) - offset;
    var itemYear = Math.floor(absoluteMonth / 12);
    var itemMonth = absoluteMonth % 12 + 1;
    var itemYearMonth = itemYear + '-' + ('0' + itemMonth).slice(-2);
    var itemLastDay = new Date(itemYear, itemMonth, 0).getDate();
    var itemEndInfo = wmReportPeriodInfo_(itemYearMonth + '-' + ('0' + itemLastDay).slice(-2));
    var itemLastWeek = Math.max(1, Math.min(5, Number(itemEndInfo && itemEndInfo.lastWeek || 1)));
    var isCurrentItem = offset === 0;
    var previousMonthBeforeMonthlyIssue = offset === 1 && day === 1 && nowTime < '00:01';
    var selectedWeek = isCurrentItem ? latestReleasedWeek : itemLastWeek;
    var reportType = isCurrentItem ? 'Weekly' : '확정 Monthly';

    /* 매월 1일 00:01 전 1분 동안에는 직전 월 Monthly가 아직 발행 전이므로
       직전 월의 마지막 일반 Weekly를 최근 성적표로 유지합니다. */
    if (previousMonthBeforeMonthlyIssue) {
      selectedWeek = Math.max(1, itemLastWeek - 1);
      reportType = 'Weekly';
    }

    months.push({
      yearMonth:itemYearMonth,
      label:itemYear + '년 ' + itemMonth + '월',
      isCurrent:isCurrentItem,
      selectedWeek:selectedWeek,
      reportType:reportType
    });
  }

  return {success:true,message:'발행 완료 최근성적표 4개월 목록 조회 완료',defaultYearMonth:months[0].yearMonth,months:months};
}

function wmMonthlyReportCompleted_(row) {
  var status = String(wmReportCell_(row, ['완료상태','완료']) || '').trim();
  if (!status || /미완료|incomplete|fail|중단|학습중/i.test(status)) return false;
  return /완료|complete|completed|success|1회차|2회차|3회차/i.test(status);
}

function wmMonthlyReportWeek_(dateText) {
  return wmReportCalendarWeek_(dateText);
}

function wmMonthlyReportRound_(row) {
  var round = wmParseOfficialCompleteRound_(wmReportCell_(row, ['완료','완료회차','학습회차','회차']));
  return round > 0 ? round : (wmMonthlyReportCompleted_(row) ? 1 : 0);
}

function wmMonthlyReportAverage_(sum, count) {
  return count ? Math.round((sum / count) * 10) / 10 : 0;
}

function wmMonthlyReportDuration_(seconds) {
  var total = Math.max(0, Math.round(Number(seconds || 0)));
  return {
    seconds:total,
    hours:Math.floor(total / 3600),
    minutesRemainder:Math.floor((total % 3600) / 60),
    totalMinutes:Math.floor(total / 60),
    text:wmReportFormatDuration_(total)
  };
}

function wmMonthlyReportDevice_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  if (/tablet|ipad|tab|tap/i.test(text)) return 'Tab';
  if (/mobile|phone|iphone|android/i.test(text)) return 'Phone';
  return 'PC';
}

function wmMonthlyReportYearMonth_(e, reportRows, recordRows) {
  var requested = wmMonthlyReportParam_(e, ['yearMonth','reportMonth','month','연월']);
  var match = requested.match(/^(\d{4})[-.\/]?(\d{1,2})$/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2);
  requested = wmMonthlyReportParam_(e, ['reportWeek','리포트주차']);
  match = requested.match(/^(\d{4})[-.\/](\d{1,2})/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2);
  var latest = '';
  (reportRows || []).concat(recordRows || []).forEach(function(row) {
    var text = String(row['리포트주차'] || wmReportNormalizeDate_(wmReportCell_(row, ['학습날짜','학습일','날짜','등록일'])) || '');
    var found = text.match(/^(\d{4})[-.\/](\d{1,2})/);
    var ym = found ? found[1] + '-' + ('0' + found[2]).slice(-2) : '';
    if (ym && (!latest || ym > latest)) latest = ym;
  });
  return latest;
}

function wmMonthlyReportLatestProgress_(rows, studentId) {
  var found = (rows || []).filter(function(row) {
    return String(wmReportCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase() === studentId;
  });
  found.sort(function(a, b) {
    return String(wmReportCell_(b, ['최종수정일','수정일','등록일','학습날짜']) || '')
      .localeCompare(String(wmReportCell_(a, ['최종수정일','수정일','등록일','학습날짜']) || ''));
  });
  return found[0] || {};
}

function wmMonthlyReportShortDate_(dateText) {
  var match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? match[1].slice(-2) + '.' + match[2] + '.' + match[3] : String(dateText || '');
}

function wmMonthlyReportDateMinute_(dateText) {
  var raw = String(dateText || '').trim();
  var match = raw.match(/(\d{2,4})\s*[-.\/]\s*(\d{1,2})\s*[-.\/]\s*(\d{1,2})(?:[ T]+(?:(오전|오후|AM|PM)\s*)?(\d{1,2}):(\d{2})(?::\d{2})?)?/i);
  if (!match) return raw.replace(/:(\d{2})(?=\s*$)/, '');
  var year = String(match[1]);
  var hour = Number(match[5] || 0);
  var marker = String(match[4] || '').toUpperCase();
  if ((marker === '오후' || marker === 'PM') && hour < 12) hour += 12;
  if ((marker === '오전' || marker === 'AM') && hour === 12) hour = 0;
  return year.slice(-2) + '.' + ('0' + match[2]).slice(-2) + '.' + ('0' + match[3]).slice(-2) +
    ' ' + ('0' + hour).slice(-2) + ':' + ('0' + (match[6] || 0)).slice(-2);
}

function wmMonthlyReportPeriod_(startDate, endDate) {
  var start = String(startDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var end = String(endDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!start) return '-';
  var startText = start[1].slice(-2) + '.' + start[2] + '.' + start[3];
  if (!end || startDate === endDate) return startText;
  return startText + '-' + end[2] + '.' + end[3];
}

function wmMonthlyReportWeekRow_(week, yearMonth) {
  var range = wmReportCalendarWeekRange_(yearMonth, week);
  var period = wmMonthlyReportPeriod_(range.startDate, range.endDate);
  return {week:week,label:week+'주차',startDate:range.startDate,endDate:range.endDate,period:period,completedSets:0,wordCount:0,scoreSum:0,scoreCount:0,averageScore:0,totalSeconds:0,totalTime:'0초',dates:{},attendanceDays:0};
}

/* =========================================================
 * [성적표보기·최근성적표보기 공통 월간·주간 성적등급 산정 기준]
 *
 * 1. Weekly Report
 * - 기존 주간 성적등급 계산식을 그대로 적용합니다.
 *
 * 2. Monthly Report 계산기간
 * - 학습시작일과 등록일이 같으면 해당 날짜를 적용합니다.
 * - 두 날짜가 다르면 학습시작일을 우선 적용합니다.
 * - 학습시작일이 없으면 등록일을 적용합니다.
 * - 최초 학습월은 학습시작일부터 말일까지 계산합니다.
 * - 중간 학습월은 매월 1일부터 말일까지 계산합니다.
 * - 마지막 학습월은 1일부터 학습종료일까지 계산합니다.
 *
 * 3. 월간결산 방식
 * - 해당 월의 달력상 4주 또는 5주 전체를 하나의 Monthly Report로 합산합니다.
 * - 달력상 6주차가 발생하면 5주차에 통합합니다.
 * - 시작일 또는 종료일 때문에 일부만 포함되는 주차는 실제 적용일수만큼
 *   출석약속과 세트약속을 일할 계산합니다.
 * - 계산기간에 전부 포함되는 주차는 주간 약속량 전체를 적용합니다.
 * - 계산기간에 포함되지만 학습을 통째로 빠진 주차는 출석 0·세트 0으로 계산합니다.
 * - 각 주차의 초과 출석·초과 세트는 다른 주차의 결석을 상쇄하지 못합니다.
 * - 주차별 성적점수를 평균하는 방식이 아니라, 주차별 인정 출석·인정 세트와
 *   약속량을 월 전체로 합산하여 월간 출석점수를 계산합니다.
 * - 학업집중도와 점수추이는 해당 월 계산기간의 실제 완료 기록 전체로 계산합니다.
 *
 * 4. 최종 등급
 * - 월간 출석점수·학업집중도·점수추이를 합산한 뒤 최종점수를 사사오입하여
 *   기존 Grade/Class 등급표에 따라 월간 성적등급을 산정합니다.
 * - 성적표 하단 좌측 집계표와 우측 자동코멘트는 반드시 같은 산정값을 사용합니다.
 * - 일반 주차는 발행 시점에 1주차부터 해당 종료 주차까지 누적한 Weekly Report로 확정 저장합니다.
 * - 월 마지막 주는 별도 Weekly를 만들지 않고 마지막 주 기록을 포함한 월 전체를 다음 달 1일 00:01 Monthly Report로 확정 발행합니다.
 * ========================================================= */
function wmMonthlySettlementLearningStart_(student) {
  return wmReportNormalizeDate_(wmReportCell_(student || {}, ['학습시작일'])) ||
    wmReportNormalizeDate_(wmReportCell_(student || {}, ['등록일']));
}

function wmMonthlySettlementLearningEnd_(student) {
  return wmReportNormalizeDate_(wmReportCell_(student || {}, ['학습종료일']));
}

function wmBuildGradeAutoComment_(data) {
  data = data || {};
  var monthly = /월간|monthly/i.test(String(data.reportType || ''));
  var reportTitle = monthly
    ? data.year + '년 ' + data.month + '월'
    : data.year + '년 ' + data.month + '월 ' + data.week + '주';
  var currentPeriod = monthly ? '이번 달' : '이번 주';
  var nextPeriod = monthly ? '다음 달' : '다음 주';
  var goal = '';
  if (Number(data.grade) === 1) goal = nextPeriod + '에도 현재 학습 리듬을 유지하며 최고 등급에 도전해 보세요.';
  else if (Number(data.grade) === 2) goal = nextPeriod + '에는 출석과 세트학습을 꾸준히 실천하여 1등급에 도전해 보세요.';
  else if (Number(data.grade) === 3) goal = nextPeriod + '에는 약속한 학습계획을 꾸준히 실천하여 더 높은 등급에 도전해 보세요.';
  return [
    data.studentName + ' 학생의 ' + reportTitle + ' 학습결과 보고서입니다.',
    data.studentName + ' 학생의 학습약속은 출석 주 ' + data.attendancePromise + '회, 세트 ' + data.setPromise + '세트로',
    '학습 결과는 출석점수 ' + data.attendanceScore + '점 · 학업집중도 ' + data.participationScore + '점 · 점수추이 ' + data.scoreTrend + '점입니다.',
    '이에 따라 ' + currentPeriod + ' 성적은 최종점수 ' + data.finalScore + '점(' + data.finalScore + '/30), 등급은 ' + data.grade + '.' + data.gradeClass + '등급입니다.',
    goal
  ].filter(function(line){ return String(line || '').trim() !== ''; }).join('\n');
}

function wmMonthlySettlementUtcDay_(isoDate) {
  var match = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])) / 86400000 : NaN;
}

function wmMonthlySettlement_(yearMonth, lastWeek, learningStartDate, learningEndDate, attendancePromise, setPromise, weeks) {
  var parts = String(yearMonth || '').match(/^(\d{4})-(\d{2})$/);
  if (!parts) return {expectedAttendance:0,expectedSets:0,creditedAttendance:0,creditedSets:0,attendanceRate:'',setProgressRate:''};
  var monthStart = yearMonth + '-01';
  var monthLastDay = new Date(Number(parts[1]), Number(parts[2]), 0).getDate();
  var monthEnd = yearMonth + '-' + ('0' + monthLastDay).slice(-2);
  var start = wmReportNormalizeDate_(learningStartDate) || monthStart;
  if (start < monthStart) start = monthStart;
  if (start > monthEnd) return {expectedAttendance:0,expectedSets:0,creditedAttendance:0,creditedSets:0,attendanceRate:'',setProgressRate:''};
  var end = wmReportNormalizeDate_(learningEndDate) || monthEnd;
  if (end > monthEnd) end = monthEnd;
  if (end < monthStart || end < start) return {expectedAttendance:0,expectedSets:0,creditedAttendance:0,creditedSets:0,attendanceRate:'',setProgressRate:''};
  var expectedAttendance = 0, expectedSets = 0, creditedAttendance = 0, creditedSets = 0;
  var weekLimit = Math.max(1, Math.min(5, Number(lastWeek || 1)));
  for (var weekNo = 1; weekNo <= weekLimit; weekNo++) {
    var range = wmReportCalendarWeekRange_(yearMonth, weekNo);
    var rangeStart = String(range.startDate || '');
    var rangeEnd = String(range.endDate || '');
    if (!rangeStart || !rangeEnd || start > rangeEnd || end < rangeStart) continue;
    var effectiveStart = start > rangeStart ? start : rangeStart;
    var effectiveEnd = end < rangeEnd ? end : rangeEnd;
    var totalDays = wmMonthlySettlementUtcDay_(rangeEnd) - wmMonthlySettlementUtcDay_(rangeStart) + 1;
    var eligibleDays = wmMonthlySettlementUtcDay_(effectiveEnd) - wmMonthlySettlementUtcDay_(effectiveStart) + 1;
    if (!isFinite(totalDays) || totalDays <= 0 || !isFinite(eligibleDays) || eligibleDays <= 0) continue;
    var fraction = eligibleDays / totalDays;
    var weekAttendanceTarget = Number(attendancePromise || 0) * fraction;
    var weekSetTarget = Number(setPromise || 0) * fraction;
    var weekData = weeks && weeks[weekNo - 1] ? weeks[weekNo - 1] : {};
    expectedAttendance += weekAttendanceTarget;
    expectedSets += weekSetTarget;
    creditedAttendance += Math.min(Number(weekData.attendanceDays || 0), weekAttendanceTarget);
    creditedSets += Math.min(Number(weekData.completedSets || 0), weekSetTarget);
  }
  return {
    expectedAttendance:expectedAttendance, expectedSets:expectedSets,
    creditedAttendance:creditedAttendance, creditedSets:creditedSets,
    attendanceRate:expectedAttendance > 0 ? Math.round(creditedAttendance / expectedAttendance * 1000) / 10 : '',
    setProgressRate:expectedSets > 0 ? Math.round(creditedSets / expectedSets * 1000) / 10 : ''
  };
}

/* 성적표보기·최근성적표보기·6.성적등급_DB가 함께 사용하는 단일 월 누적 성적등급 계산함수입니다. */
function wmCalculateMonthlyCumulativeGrade_(data) {
  data = data || {};
  var round1 = function(value){ return Math.round(Number(value || 0) * 10) / 10; };
  var completedSets = Number(data.completedSets || 0);
  var totalSeconds = Number(data.totalSeconds || 0);
  var scoreSum = Number(data.scoreSum || 0);
  var scoreCount = Number(data.scoreCount || 0);
  var settlement = wmMonthlySettlement_(
    data.yearMonth,
    data.selectedWeek,
    data.learningStartDate,
    data.learningEndDate,
    data.attendancePromise,
    data.setPromise,
    data.weeks
  );
  var attendanceRate = settlement.attendanceRate;
  var setProgressRate = settlement.setProgressRate;
  var attendanceScore = attendanceRate !== '' && setProgressRate !== ''
    ? (setProgressRate >= 100 ? 10 : round1(Math.min(10, ((setProgressRate * 0.8) + (attendanceRate * 0.2)) / 10)))
    : '';
  var averageSetSeconds = completedSets > 0 ? Math.round(totalSeconds / completedSets) : 0;
  var averageSetMinutes = averageSetSeconds / 60;
  var participationScore = completedSets <= 0 ? ''
    : (averageSetMinutes <= 20 ? 10 : averageSetMinutes <= 30 ? 9 : averageSetMinutes <= 35 ? 8 : averageSetMinutes <= 40 ? 7 : averageSetMinutes <= 45 ? 6 : 5);
  var averageRawScore = scoreCount ? round1(scoreSum / scoreCount) : '';
  var convertedRawScore = averageRawScore !== '' ? round1(Math.max(0, Math.min(10, averageRawScore / 10))) : '';
  var standardScore = completedSets <= 0 ? ''
    : (averageSetMinutes <= 30 ? 10 : averageSetMinutes <= 40 ? 9 : averageSetMinutes <= 50 ? 8 : 7);
  var scoreTrend = convertedRawScore !== '' && standardScore !== '' ? round1((convertedRawScore * 0.9) + (standardScore * 0.1)) : '';
  var finalScore = attendanceScore !== '' && participationScore !== '' && scoreTrend !== ''
    ? round1(attendanceScore + participationScore + scoreTrend)
    : '';
  var grade = '', gradeClass = '';
  if (finalScore !== '') {
    var roundedFinalScore = Math.round(finalScore);
    if (roundedFinalScore >= 1 && roundedFinalScore <= 30) {
      grade = Math.floor((30 - roundedFinalScore) / 10) + 1;
      gradeClass = ((30 - roundedFinalScore) % 10) + 1;
    }
  }
  return {
    settlement:settlement,
    attendanceRate:attendanceRate,
    setProgressRate:setProgressRate,
    attendanceScore:attendanceScore,
    averageSetSeconds:averageSetSeconds,
    participationScore:participationScore,
    averageRawScore:averageRawScore,
    convertedRawScore:convertedRawScore,
    standardScore:standardScore,
    scoreTrend:scoreTrend,
    finalScore:finalScore,
    grade:grade,
    gradeClass:gradeClass
  };
}

function wmGetMonthlyReportDataForLms_(e) {
  try {
    var studentId = wmMonthlyReportParam_(e, ['studentId','studentID','sid','학생ID']).toUpperCase();
    if (!studentId) return {success:false,message:'학생ID가 없습니다.',data:null};
    var reportSource = wmMonthlyReportParam_(e, ['requestSource','source']).toLowerCase();
    if (reportSource === 'map') {
      var mapSessionToken = wmMonthlyReportParam_(e, ['sessionToken','wmSessionToken','현재세션']);
      if (!mapSessionToken || !wmIsStudentSessionValid_(studentId, mapSessionToken)) {
        return {success:false,sessionExpired:true,message:'학생 세션을 확인할 수 없습니다.',data:null};
      }
    }

    var studentsResult = wmReadLmsSheetAsObjects_('1.학생관리_DB');
    var recordsResult = wmReadLmsSheetAsObjects_('2.학습기록_DB');
    var reportsResult = wmReadLmsSheetAsObjects_('4-1.성적표생성_DB');
    var progressResult = wmReadLmsSheetAsObjects_('8.현재진행_DB');
    var gradeResult = wmReadLmsSheetAsObjects_('6.성적등급_DB');
    var results = [studentsResult,recordsResult,reportsResult,progressResult,gradeResult];
    for (var x = 0; x < results.length; x++) {
      if (!results[x].success) return {success:false,message:results[x].message || '성적표 DB 조회 실패',data:null};
    }

    var students = studentsResult.rows || [];

    var actor = wmGetLmsRequestActor_(e);
    if (reportSource !== 'map' && reportSource !== 'publicreport' && actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      var allowed = wmStudentIdSet_(wmFilterStudentsForActor_(students, actor));
      if (!allowed[studentId]) return {success:false,message:'조회 권한이 없습니다.',data:null};
    }

    var student = null;
    for (var s = 0; s < students.length; s++) {
      if (String(wmReportCell_(students[s], ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase() === studentId) {
        student = students[s];
        break;
      }
    }
    if (!student) return {success:false,message:'1.학생관리_DB에서 학생ID를 찾을 수 없습니다: '+studentId,data:null};

    var reports = (reportsResult.rows || []).filter(function(row) {
      return String(wmReportCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase() === studentId;
    });
    var records = (recordsResult.rows || []).filter(function(row) {
      return String(wmReportCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase() === studentId;
    });
    var yearMonth = wmMonthlyReportYearMonth_(e, reports, records);
    if (!yearMonth) return {success:false,message:'조회할 연월을 확인할 수 없습니다.',data:null};
    reports = reports.filter(function(row) {
      var week = String(row['리포트주차'] || '');
      var dateText = wmReportNormalizeDate_(row['학습날짜']);
      return week.indexOf(yearMonth + '-') === 0 || dateText.slice(0,7) === yearMonth;
    }).sort(function(a,b) {
      return String(b['리포트주차'] || b['학습날짜'] || '').localeCompare(String(a['리포트주차'] || a['학습날짜'] || ''));
    });

    var requestedWeek = wmMonthlyReportParam_(e, ['reportWeek','리포트주차']);
    var weekMatch = requestedWeek.match(/(\d+)주차/);
    var hasRequestedWeek = !!weekMatch;

    /* WM_MONTHLY_REPORT_RELEASE_UNIFIED_V1
     * 화면 주차는 달력 날짜가 아니라 실제 발행이 완료된 6.성적등급_DB 행을 기준으로 확정합니다.
     * 일반 주차는 주간 행, 마지막 주차는 월간 행이 실제 생성된 경우에만 노출합니다.
     * 따라서 상단 제목·하단 점수표·자동코멘트가 서로 다른 주차를 가리키지 않습니다.
     */
    var selectedMonthParts = yearMonth.match(/^(\d{4})-(\d{2})$/);
    if (!selectedMonthParts) return {success:false,message:'조회할 연월을 확인할 수 없습니다.',data:null};
    var selectedMonthLastDay = new Date(Number(selectedMonthParts[1]), Number(selectedMonthParts[2]), 0).getDate();
    var selectedMonthEndDate = yearMonth + '-' + ('0' + selectedMonthLastDay).slice(-2);
    var selectedMonthEndInfo = wmReportPeriodInfo_(selectedMonthEndDate);
    var selectedMonthLastWeek = Math.max(1, Math.min(5, Number(selectedMonthEndInfo && (selectedMonthEndInfo.lastWeek || selectedMonthEndInfo.week) || 1)));
    var releasedGradeRows = (gradeResult.rows || []).filter(function(row) {
      var gradeStudentId = String(wmReportCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase();
      var gradeYear = wmReportParseNumber_(wmReportCell_(row, ['연도','Year','year']));
      var gradeMonth = wmReportParseNumber_(wmReportCell_(row, ['월','Month','month']));
      var gradeWeek = wmReportParseNumber_(wmReportCell_(row, ['주차','리포트주차','Week']));
      var gradeType = String(wmReportCell_(row, ['리포트유형','유형','Report_Type','reportType']) || '').trim();
      var gradePublishStatus = String(wmReportCell_(row, ['성적표반영상태']) || '').trim();
      var isMonthlyRow = /월간|monthly/i.test(gradeType);
      if (gradeStudentId !== studentId || gradeYear !== Number(selectedMonthParts[1]) || gradeMonth !== Number(selectedMonthParts[2])) return false;
      if (gradePublishStatus !== '반영완료') return false;
      if (gradeWeek < 1 || gradeWeek > selectedMonthLastWeek) return false;
      return gradeWeek === selectedMonthLastWeek ? isMonthlyRow : !isMonthlyRow;
    }).sort(function(a,b) {
      var aWeek = wmReportParseNumber_(wmReportCell_(a, ['주차','리포트주차','Week']));
      var bWeek = wmReportParseNumber_(wmReportCell_(b, ['주차','리포트주차','Week']));
      if (aWeek !== bWeek) return bWeek - aWeek;
      return String(wmReportCell_(b, ['생성일시','수정일','등록일']) || '').localeCompare(String(wmReportCell_(a, ['생성일시','수정일','등록일']) || ''));
    });
    if (!releasedGradeRows.length) return {success:false,message:'아직 발행된 성적표가 없습니다.',data:null};

    var selectedReleaseGradeRow = releasedGradeRows[0];
    var selectedWeek = wmReportParseNumber_(wmReportCell_(selectedReleaseGradeRow, ['주차','리포트주차','Week']));
    if (hasRequestedWeek) {
      var requestedWeekNumber = Math.max(1, Math.min(5, Number(weekMatch[1])));
      var requestedReleaseGradeRow = null;
      for (var releasedIndex = 0; releasedIndex < releasedGradeRows.length; releasedIndex++) {
        if (wmReportParseNumber_(wmReportCell_(releasedGradeRows[releasedIndex], ['주차','리포트주차','Week'])) === requestedWeekNumber) {
          requestedReleaseGradeRow = releasedGradeRows[releasedIndex];
          break;
        }
      }
      if (!requestedReleaseGradeRow) return {success:false,message:'아직 발행되지 않은 성적표입니다.',data:null};
      selectedReleaseGradeRow = requestedReleaseGradeRow;
      selectedWeek = requestedWeekNumber;
    }
    var isClosedMonthly = selectedWeek === selectedMonthLastWeek && /월간|monthly/i.test(String(wmReportCell_(selectedReleaseGradeRow, ['리포트유형','유형','Report_Type','reportType']) || ''));
    var selectedReportRow = wmReportSelectExactRow_(reports, studentId, yearMonth, selectedWeek, isClosedMonthly ? '월간' : '주간');
    if (!selectedReportRow) return {success:false,message:'선택한 발행 성적표 행을 찾을 수 없습니다.',data:null};
    var learningStartDate = wmReportNormalizeDate_(wmReportCell_(selectedReleaseGradeRow, ['학습시작일']));
    var learningEndDate = wmReportNormalizeDate_(wmReportCell_(selectedReleaseGradeRow, ['학습종료일']));


    var progress = wmMonthlyReportLatestProgress_(progressResult.rows || [], studentId);
    var currentLevel = Number(String(progress['현재레벨'] || progress['Set_ID'] || '').match(/\d+/) || 0);
    var levelLabel = currentLevel ? 'Basic '+currentLevel : wmReportLevelLabel_(selectedReportRow['Set_ID'], wmReportCell_(student, ['최초배정','학습배정','학습레벨']));
    var levelCompletionCondition = wmReportParseNumber_(progress['레벨완료조건'] || wmReportCell_(student, ['레벨완료조건'])) || 1;
    var storedLevelCompletionCount = wmReportParseNumber_(progress['레벨완료횟수']);
    var lastRound = wmReportParseNumber_(progress['순차완주회차'] || progress['완료회차']);
    if (!lastRound) {
      lastRound = storedLevelCompletionCount >= levelCompletionCondition
        ? levelCompletionCondition
        : storedLevelCompletionCount + 1;
    }
    lastRound = Math.max(1, Math.min(levelCompletionCondition, lastRound));

    var allCompleted = records.filter(function(row) {
      return wmMonthlyReportCompleted_(row);
    }).map(function(row) {
      var item = {}, key;
      for (key in row) if (row.hasOwnProperty(key)) item[key] = row[key];
      item.__dateTime = String(wmReportCell_(row, ['학습날짜','학습일','날짜','등록일']) || '').trim();
      item.__date = wmReportNormalizeDate_(item.__dateTime);
      item.__setId = wmReportPlainSetId_(wmReportCell_(row, ['Set_ID','Set ID','세트ID']));
      item.__level = wmReportLevelFromSet_(item.__setId);
      item.__round = wmMonthlyReportRound_(row);
      item.__scoreText = wmReportCell_(row, ['점수','한영주관식']);
      item.__score = wmReportParseNumber_(item.__scoreText);
      item.__seconds = wmReportParseSeconds_(wmReportCell_(row, ['총소요시간','총 소요시간']));
      return item;
    }).sort(function(a,b) {
      return String(a.__dateTime || a.__date).localeCompare(String(b.__dateTime || b.__date)) || String(a['학습기록ID'] || '').localeCompare(String(b['학습기록ID'] || ''));
    });
    var completed = allCompleted.filter(function(row) {
      var recordWeek = wmMonthlyReportWeek_(row.__date);
      return row.__date.slice(0,7) === yearMonth
        && (!learningStartDate || row.__date >= learningStartDate)
        && (!learningEndDate || row.__date <= learningEndDate)
        && (hasRequestedWeek ? recordWeek <= selectedWeek : (!selectedWeek || recordWeek <= selectedWeek));
    });

    if (!hasRequestedWeek && !selectedWeek && completed.length) selectedWeek = wmMonthlyReportWeek_(completed[completed.length - 1].__date);
    if (!hasRequestedWeek && !selectedWeek) selectedWeek = 1;

    var requestTimeZone = Session.getScriptTimeZone() || 'Asia/Seoul';
    var requestCurrentMonth = Utilities.formatDate(new Date(), requestTimeZone, 'yyyy-MM');
    var isCurrentReportMonth = yearMonth === requestCurrentMonth;
    var periodEndRange = wmReportCalendarWeekRange_(yearMonth, Math.max(1, selectedWeek || 1));
    var periodEndDate = periodEndRange.endDate || yearMonth + '-01';
    if (learningEndDate && learningEndDate < periodEndDate) periodEndDate = learningEndDate;
    var periodStartDate = yearMonth + '-01';
    if (learningStartDate && learningStartDate > periodStartDate) periodStartDate = learningStartDate;
    if (periodStartDate > periodEndDate) return {success:false,message:'성적 내용이 없습니다.',data:null};
    var periodHistory = allCompleted.filter(function(row) { return row.__date && row.__date <= periodEndDate; });
    var periodLastRecord = periodHistory.length ? periodHistory[periodHistory.length - 1] : null;

    if (!isCurrentReportMonth && periodLastRecord) {
      currentLevel = periodLastRecord.__level;
      levelLabel = currentLevel ? 'Basic ' + currentLevel : levelLabel;
      lastRound = Math.max(1, periodLastRecord.__round || 1);
      levelCompletionCondition = wmReportParseNumber_(selectedReportRow['레벨완료조건'] || wmReportCell_(student, ['레벨완료조건'])) || levelCompletionCondition;
    }

    var weeks = [
      wmMonthlyReportWeekRow_(1,yearMonth),
      wmMonthlyReportWeekRow_(2,yearMonth),
      wmMonthlyReportWeekRow_(3,yearMonth),
      wmMonthlyReportWeekRow_(4,yearMonth),
      wmMonthlyReportWeekRow_(5,yearMonth)
    ];
    var monthly = {completedSets:0,wordCount:0,scoreSum:0,scoreCount:0,averageScore:0,totalSeconds:0,dates:{},attendanceDays:0,calculationStartDate:periodStartDate,calculationEndDate:periodEndDate};
    completed.forEach(function(row) {
      var weekNo = wmMonthlyReportWeek_(row.__date);
      if (!weekNo) return;
      var week = weeks[weekNo-1];
      var words = wmReportWordCountForSet_(row.__setId);
      week.completedSets++; week.wordCount += words; week.totalSeconds += row.__seconds; week.dates[row.__date] = true;
      monthly.completedSets++; monthly.wordCount += words; monthly.totalSeconds += row.__seconds; monthly.dates[row.__date] = true;
      if (String(row.__scoreText || '').trim() !== '' && isFinite(row.__score)) {
        week.scoreSum += row.__score; week.scoreCount++;
        monthly.scoreSum += row.__score; monthly.scoreCount++;
      }
    });
    weeks.forEach(function(week) {
      week.averageScore = wmMonthlyReportAverage_(week.scoreSum, week.scoreCount);
      week.totalTime = wmReportFormatDuration_(week.totalSeconds);
      week.attendanceDays = Object.keys(week.dates).length;
      delete week.scoreSum; delete week.scoreCount; delete week.dates;
    });
    monthly.averageScore = wmMonthlyReportAverage_(monthly.scoreSum, monthly.scoreCount);
    monthly.attendanceDays = Object.keys(monthly.dates).length;
    monthly.totalTime = wmMonthlyReportDuration_(monthly.totalSeconds);
    delete monthly.scoreSum; delete monthly.scoreCount; delete monthly.dates;

    var totalSets = isCurrentReportMonth ? wmReportParseNumber_(progress['전체세트수']) : 0;
    if (!totalSets && currentLevel) totalSets = getLevelSetSequence_(currentLevel).length;
    var inProgress = isCurrentReportMonth && String(progress['Set_ID'] || '').trim() && String(progress['완료Step'] || '').trim() !== 'STEP6' ? 1 : 0;

    var currentRoundCompleted = allCompleted.filter(function(row) {
      return (!currentLevel || row.__level === currentLevel) && row.__round === lastRound && (isCurrentReportMonth || row.__date <= periodEndDate);
    });
    var currentRoundSetIds = {};
    currentRoundCompleted.forEach(function(row) {
      if (row.__setId) currentRoundSetIds[row.__setId] = true;
    });
    var completedSets = Object.keys(currentRoundSetIds).length;
    var roundCompleted = !!(totalSets && completedSets >= totalSets);
    if (!isCurrentReportMonth) storedLevelCompletionCount = Math.max(0, lastRound - 1) + (roundCompleted ? 1 : 0);
    storedLevelCompletionCount = Math.min(levelCompletionCondition, storedLevelCompletionCount);
    var percentage = totalSets ? Math.round(completedSets / totalSets * 100) : 0;
    percentage = Math.max(0,Math.min(100,percentage));

    var scoreTrend = weeks.map(function(week) {
      return {week:week.week,label:week.label,averageScore:week.averageScore,completedSets:week.completedSets};
    });

    var selectedDetailRows = completed.filter(function(row) {
      return wmMonthlyReportWeek_(row.__date) <= selectedWeek;
    }).sort(function(a,b) {
      var weekDiff = wmMonthlyReportWeek_(b.__date) - wmMonthlyReportWeek_(a.__date);
      if (weekDiff) return weekDiff;
      return String(b.__dateTime || b.__date).localeCompare(String(a.__dateTime || a.__date)) || String(b['학습기록ID'] || '').localeCompare(String(a['학습기록ID'] || ''));
    }).slice(0,10);
    selectedDetailRows.sort(function(a,b) {
      var weekDiff = wmMonthlyReportWeek_(a.__date) - wmMonthlyReportWeek_(b.__date);
      if (weekDiff) return weekDiff;
      return String(a.__dateTime || a.__date).localeCompare(String(b.__dateTime || b.__date)) || String(a['학습기록ID'] || '').localeCompare(String(b['학습기록ID'] || ''));
    });
    var detailRows = selectedDetailRows.map(function(row,index) {
      var detailDate = String(row.__date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
      var detailMonth = detailDate ? Number(detailDate[2]) : Number(yearMonth.split('-')[1]);
      return {serial:index+1,reportWeek:detailMonth+'월 '+wmMonthlyReportWeek_(row.__date)+'주차',learningDate:wmMonthlyReportDateMinute_(row.__dateTime),learningLevel:wmReportLevelLabel_(row.__setId,''),setId:row.__setId,setLearningTime:wmReportFormatDuration_(row.__seconds),wordCount:wmReportWordCountForSet_(row.__setId),score:String(row.__scoreText || '').trim()===''?'':row.__score,testTime:wmReportFormatDuration_(wmReportParseSeconds_(row['Test_총시간'])),device:wmMonthlyReportDevice_(row['기기정보'])};
    });
    while (detailRows.length < 10) detailRows.push({serial:detailRows.length+1,reportWeek:'-',learningDate:'-',learningLevel:'-',setId:'-',setLearningTime:'-',wordCount:'-',score:'-',testTime:'-',device:'-'});

    var ym = yearMonth.split('-');
    var selectedGradeRows = (gradeResult.rows || []).filter(function(row) {
      var gradeStudentId = String(wmReportCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase();
      var gradeYear = wmReportParseNumber_(wmReportCell_(row, ['연도','Year','year']));
      var gradeMonth = wmReportParseNumber_(wmReportCell_(row, ['월','Month','month']));
      return gradeStudentId === studentId && gradeYear === Number(ym[0]) && gradeMonth === Number(ym[1]);
    });
    /* 발행 시점에 6.성적등급_DB에 확정 저장된 한 행만 좌측표·우측코멘트의 유일한 원본으로 사용합니다.
       성적표를 여는 시점에는 2.학습기록_DB로 성적등급을 다시 계산하지 않습니다. */
    var selectedGradeDbRow = selectedReleaseGradeRow || null;
    var selectedGrade = null;
    if (selectedGradeDbRow) {
      selectedGrade = {};
      Object.keys(selectedGradeDbRow).forEach(function(key) { selectedGrade[key] = selectedGradeDbRow[key]; });
    }
    return {success:true,message:'발행 확정 성적표 조회 완료',data:{
      query:{studentId:studentId,yearMonth:yearMonth,selectedWeek:selectedWeek},
      period:{year:Number(ym[0]),month:Number(ym[1]),week:selectedWeek,label:isClosedMonthly?'Monthly':selectedWeek+'주차',reportType:isClosedMonthly?'월간':'주간',isClosedMonthly:isClosedMonthly},
      student:{school:wmReportCell_(student,['학교','학교명'])||selectedReportRow['학교']||'',grade:wmReportCell_(student,['학년'])||selectedReportRow['학년']||'',studentName:wmReportCell_(student,['학생이름','학생명'])||selectedReportRow['학생이름']||'',studentId:studentId,className:wmReportCell_(student,['반명','Class','학급명','클래스명'])||'',learningLevel:levelLabel},
      monthly:monthly,
      weekly:weeks,
      scoreTrend:scoreTrend,
      levelProgress:{level:levelLabel,percentage:percentage,levelCompletionCondition:levelCompletionCondition,levelCompletionCount:storedLevelCompletionCount,currentRound:lastRound,currentRoundStatus:lastRound+'회차 '+(roundCompleted?'완료':'진행 중'),totalSets:totalSets,completedSets:completedSets,inProgressSets:inProgress,remainingSets:Math.max(0,totalSets-completedSets)},
      detailRows:detailRows,
      comments:{teacherComment:String(selectedReportRow['Teacher’s Comment']||''),encouragementMessage:String(selectedReportRow['격려메시지']||''),logoUrl:String(selectedReportRow['로고URL']||'')},
      grade:{found:!!selectedGrade,gradeComment:selectedGrade,guide:wmGradeGuideText_(),commentSource:'6.성적등급_DB 발행확정행'},
      meta:{sourceRecordCount:records.length,completedMonthlyRecordCount:completed.length,currentRoundRecordCount:currentRoundCompleted.length,currentRoundUniqueSetCount:completedSets,reportRowCount:reports.length,gradeRowCount:selectedGradeRows.length,hasStoredMonthlyGrade:!!selectedGradeDbRow,learningStartDate:learningStartDate,learningEndDate:learningEndDate,calendarWeekPolicy:'월요일~일요일, 월별 기록 분리, 6주차는 5주차에 병합',reportRound:lastRound,monthlyPolicy:'학습시작일 우선·없으면 등록일, 첫 달은 시작일부터 말일까지·중간 달은 1일부터 말일까지·마지막 달은 1일부터 학습종료일까지 월간결산, 주차별 약속 초과분의 결석 상쇄 금지',scoreGraphPolicy:'선택 월의 완료 기록 전체를 주차별 평균점수로 연결',levelProgressPolicy:'해당 기간 말일 기준 마지막 학습 레벨·회차 반영',detailPolicy:'선택 월 완료 기록 최신 10건 선정 후 주차·날짜·세트 오름차순 표시, 학습일시는 분까지 표시',completedPolicy:'완료상태 또는 완료 컬럼에 완료값이 있는 기록만 포함'}
    }};
  } catch (err) {
    return {success:false,message:'월간 성적표 통합 조회·계산 오류',error:String(err && err.message ? err.message : err),data:null};
  }
}

/* WM_LMS_CUMULATIVE_WRONG_V1
 * LMS 화면 전용 가상 컬럼입니다. 2.학습기록_DB에 '누적오답' 컬럼을 만들지 않습니다.
 * 학생ID의 전체 학습기록에서 최신 학습레벨을 찾고, 그 레벨의 틀린단어만 Set_ID+단어 기준으로 중복 제거합니다.
 * 초기 학습기록 조회와 분리하여 [누적오답] 클릭 시에만 실행합니다.
 */
function wmParseWrongWordsCellForLmsCumulative_(value) {
  var parsed = value;
  if (typeof parsed === 'string') {
    var text = String(parsed || '').trim();
    if (!text) return {available:false, engKoWrite:[], koEngWrite:[]};
    try { parsed = JSON.parse(text); }
    catch (err) { return {available:false, engKoWrite:[], koEngWrite:[]}; }
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return {available:false, engKoWrite:[], koEngWrite:[]};
  }
  var available = String(parsed.version || '') === 'WM_WRONG_WORDS_V1' ||
    Array.isArray(parsed.engKoWrite) || Array.isArray(parsed.koEngWrite);
  if (!available) return {available:false, engKoWrite:[], koEngWrite:[]};
  return {
    available:true,
    engKoWrite:wmNormalizeWrongWordListForRecord_(parsed.engKoWrite),
    koEngWrite:wmNormalizeWrongWordListForRecord_(parsed.koEngWrite)
  };
}

function wmCumulativeWrongLevelFromSetId_(setId) {
  var parsed = parseWordMateSetId(String(setId || '').trim());
  return parsed && Number(parsed.level || 0) ? Number(parsed.level) : 0;
}

function wmBuildCumulativeWrongForLmsStudent_(studentId) {
  var sid = String(studentId || '').trim().toUpperCase();
  if (!sid) {
    return {success:false, available:false, message:'학생ID가 없습니다.', studentId:'', level:0, engKoWrite:[], koEngWrite:[]};
  }

  var cacheKey = wmCacheKey_('LMS_CUMULATIVE_WRONG', sid);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && cached.studentId === sid) return cached;

  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('2.학습기록_DB');
  if (!sheet) {
    return {success:false, available:false, message:'2.학습기록_DB 시트를 찾을 수 없습니다.', studentId:sid, level:0, engKoWrite:[], koEngWrite:[]};
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return {success:true, available:false, message:'학습기록이 없습니다.', studentId:sid, level:0, engKoWrite:[], koEngWrite:[]};
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(value){
    return String(value || '').trim();
  });
  var idxStudent = headers.indexOf('학생ID');
  if (idxStudent < 0) {
    return {success:false, available:false, message:'2.학습기록_DB에서 학생ID 컬럼을 찾을 수 없습니다.', studentId:sid, level:0, engKoWrite:[], koEngWrite:[]};
  }

  var rowNumbers = wmGetStudentRowNumbersFast_(sheet, headers, sid);
  if (!rowNumbers.length) {
    return {success:true, available:false, message:'해당 학생의 학습기록이 없습니다.', studentId:sid, level:0, engKoWrite:[], koEngWrite:[]};
  }

  var values = wmReadGroupedRowsFast_(sheet, rowNumbers, lastCol);
  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var row = values[r] || [];
    if (String(row[idxStudent] || '').trim().toUpperCase() !== sid) continue;
    var objectRow = {};
    for (var c = 0; c < headers.length; c++) objectRow[headers[c]] = row[c];
    rows.push(objectRow);
  }

  rows.sort(function(a, b){
    return getLearningRecordSortTime_(b['학습날짜']) - getLearningRecordSortTime_(a['학습날짜']);
  });

  var latestLevel = 0;
  for (var i = 0; i < rows.length; i++) {
    latestLevel = wmCumulativeWrongLevelFromSetId_(rows[i]['Set_ID']);
    if (latestLevel) break;
  }
  if (!latestLevel) {
    return {success:true, available:false, message:'최신 학습레벨을 확인할 수 없습니다.', studentId:sid, level:0, engKoWrite:[], koEngWrite:[]};
  }

  var result = {available:false, engKoWrite:[], koEngWrite:[]};
  var engSeen = {};
  var koSeen = {};
  rows.forEach(function(row){
    var setId = String(row['Set_ID'] || '').trim();
    if (wmCumulativeWrongLevelFromSetId_(setId) !== latestLevel) return;
    var data = wmParseWrongWordsCellForLmsCumulative_(row['틀린단어']);
    if (!data.available) return;
    result.available = true;
    data.engKoWrite.forEach(function(item){
      var key = setId.toLowerCase() + '\u0000' + String(item.word || '').toLowerCase();
      if (engSeen[key]) return;
      engSeen[key] = true;
      result.engKoWrite.push({setId:setId, word:item.word, pos:item.pos, meaning:item.meaning});
    });
    data.koEngWrite.forEach(function(item){
      var key = setId.toLowerCase() + '\u0000' + String(item.word || '').toLowerCase();
      if (koSeen[key]) return;
      koSeen[key] = true;
      result.koEngWrite.push({setId:setId, word:item.word, pos:item.pos, meaning:item.meaning});
    });
  });

  var payload = {
    success:true,
    available:result.available,
    message:result.available ? '누적오답 조회 성공' : '누적오답 데이터가 없습니다.',
    studentId:sid,
    level:latestLevel,
    title:'누적오답(영한 ' + result.engKoWrite.length + '·한영 ' + result.koEngWrite.length + ') · ' + latestLevel + '레벨',
    buttonLabel:'누적오답(영한 ' + result.engKoWrite.length + '·한영 ' + result.koEngWrite.length + ')',
    engKoWrite:result.engKoWrite,
    koEngWrite:result.koEngWrite
  };
  wmCachePutJson_(cacheKey, payload, 60);
  return payload;
}

function wmGetCumulativeWrongForLmsApi_(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var studentId = String(p.studentId || p.studentID || p.Student_ID || '').trim().toUpperCase();
    if (!studentId) {
      return outputResult(e, {success:false, available:false, message:'학생ID가 없습니다.', studentId:'', level:0, engKoWrite:[], koEngWrite:[]});
    }

    var actor = wmGetLmsRequestActor_(e);
    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      var allowed = wmAllowedStudentIdSetForActor_(actor) || {};
      if (!allowed[studentId]) {
        return outputResult(e, {success:false, available:false, message:'해당 학생의 조회 권한이 없습니다.', studentId:studentId, level:0, engKoWrite:[], koEngWrite:[]});
      }
    }

    return outputResult(e, wmBuildCumulativeWrongForLmsStudent_(studentId));
  } catch (err) {
    return outputResult(e, {success:false, available:false, message:'누적오답 조회 오류', error:String(err && err.message ? err.message : err), studentId:'', level:0, engKoWrite:[], koEngWrite:[]});
  }
}

function getLearningRecordsForLmsApi_(e) {
  try {
    var result = wmIsLmsInitialPageRequest_(e) ? wmReadLmsFirstRowsAsObjects_('2.학습기록_DB', wmLmsInitialLimit_(e, 25)) : wmReadLmsSheetAsObjects_('2.학습기록_DB');
    var actor = wmGetLmsRequestActor_(e);
    var rows = result.rows || [];
    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      rows = wmFilterRowsByStudentSet_(rows, wmAllowedStudentIdSetForActor_(actor) || {}, actor);
    }

    /* WM_PROFILE_SINGLE_STUDENT_SELECTION_V1
     * 프로필센터 후보 선택 후 선택된 학생ID 1개만 exact 기준으로 조회합니다.
     */
    /* WM_PROFILE_STUDY_RECORD_DB_LOOKUP_V2
     * 프로필센터 학습기록은 복수 학생ID와 기간을 서버에서 먼저 필터링합니다.
     * 기간이 없으면 오늘 기준 최근 1개월을 기본값으로 사용합니다.
     */
    var requestedStudentIdsText = String(
      e && e.parameter ? (
        e.parameter.studentIds ||
        e.parameter.studentId ||
        e.parameter.Student_ID ||
        ''
      ) : ''
    ).trim();

    var requestedStudentIds = {};
    requestedStudentIdsText.split(',').forEach(function(value) {
      var id = String(value || '').trim().toUpperCase();
      if (id) requestedStudentIds[id] = true;
    });

    /* WM_LMS_RECORD_SEARCH_SERVER_FILTER_V1_20260822
     * LMS 학습기록 검색은 최초 25행 캐시를 브라우저에서 재필터링하지 않고,
     * 검색 버튼을 눌렀을 때 서버에서 요청 조건으로 다시 필터링합니다.
     * 본사용 LMS 조회는 학생 Map의 3년 표시 제한과 분리하며 DB 원본 보관기간을 제한하지 않습니다. */
    var requestedStudentName = String(e && e.parameter ? (e.parameter.studentName || '') : '').trim().toLowerCase();
    var requestedTeacher = String(e && e.parameter ? (e.parameter.teacher || '') : '').trim().toLowerCase();
    var requestedLeader = String(e && e.parameter ? (e.parameter.leader || '') : '').trim().toLowerCase();
    var requestedClass = String(e && e.parameter ? (e.parameter.className || e.parameter.class || '') : '').trim().toLowerCase();
    var requestedSchool = String(e && e.parameter ? (e.parameter.school || '') : '').trim().toLowerCase();
    var requestedRegion = String(e && e.parameter ? (e.parameter.region || '') : '').trim().toLowerCase();
    var requestedLevel = String(e && e.parameter ? (e.parameter.level || '') : '').trim().toLowerCase();

    function wmProfileParseDateOnly_(value) {
      var text = String(value || '').trim();
      var match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!match) return null;
      var date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
      date.setHours(0, 0, 0, 0);
      return isNaN(date.getTime()) ? null : date;
    }

    var includeAllDates = String(
      e && e.parameter ? (e.parameter.allDates || '') : ''
    ).trim() === '1';

    var requestedStart = wmProfileParseDateOnly_(
      e && e.parameter ? e.parameter.dateStart : ''
    );
    var requestedEnd = wmProfileParseDateOnly_(
      e && e.parameter ? e.parameter.dateEnd : ''
    );

    if (!includeAllDates && !requestedStart && !requestedEnd) {
      requestedEnd = new Date();
      requestedEnd.setHours(0, 0, 0, 0);
      requestedStart = new Date(requestedEnd);
      requestedStart.setMonth(requestedStart.getMonth() - 1);
    } else if (!includeAllDates) {
      if (!requestedStart) requestedStart = new Date(requestedEnd);
      if (!requestedEnd) {
        requestedEnd = new Date();
        requestedEnd.setHours(0, 0, 0, 0);
      }
    }

    rows = rows.filter(function(row) {
      var rowId = String(row['학생ID'] || '').trim().toUpperCase();
      if (Object.keys(requestedStudentIds).length && !requestedStudentIds[rowId]) {
        return false;
      }

      if (requestedStudentName && String(row['학생이름'] || '').trim().toLowerCase().indexOf(requestedStudentName) === -1) return false;
      if (requestedTeacher && String(row['교사명'] || '').trim().toLowerCase().indexOf(requestedTeacher) === -1) return false;
      if (requestedLeader && String(row['담당리더'] || row['리더'] || row['리더명'] || '').trim().toLowerCase().indexOf(requestedLeader) === -1) return false;
      if (requestedClass && String(row['Class'] || row['반명'] || row['반'] || '').trim().toLowerCase().indexOf(requestedClass) === -1) return false;
      if (requestedSchool && String(row['학교'] || '').trim().toLowerCase().indexOf(requestedSchool) === -1) return false;
      if (requestedRegion && String(row['지역'] || row['주소'] || '').trim().toLowerCase().indexOf(requestedRegion) === -1) return false;
      if (requestedLevel) {
        var rowLevel = String(row['학습레벨'] || '').trim().toLowerCase();
        var rowSetId = String(row['Set_ID'] || '').trim().toLowerCase().replace(/^wm/, '');
        if (rowLevel.indexOf(requestedLevel) === -1 &&
            rowSetId.indexOf(requestedLevel + '-') !== 0 &&
            rowSetId.indexOf(requestedLevel) === -1) return false;
      }

      if (!includeAllDates) {
        var rowDate = wmProfileParseDateOnly_(row['학습날짜']);
        if (!rowDate) return false;
        if (requestedStart && rowDate < requestedStart) return false;
        if (requestedEnd && rowDate > requestedEnd) return false;
      }
      return true;
    });

    return outputResult(e, {
      success: !!result.success,
      message: result.message || '',
      records: rows,
      learningRecords: rows
    });
  } catch (err) {
    return outputResult(e, { success:false, message:'2.학습기록_DB 조회 오류', error:String(err && err.message ? err.message : err), records:[], learningRecords:[] });
  }
}

function wmCurrentProgressRecentSortTime_(value) {
  var text = String(value || '').trim();
  if (!text) return 0;
  var normalized = text.replace(' ', 'T');
  var parsed = new Date(normalized);
  var time = parsed.getTime();
  if (isFinite(time)) return time;
  var match = text.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?/);
  if (!match) return 0;
  return new Date(
    Number(match[1]), Number(match[2]) - 1, Number(match[3]),
    Number(match[4]), Number(match[5]), Number(match[6]),
    Number(String(match[7] || '0').padEnd(3, '0'))
  ).getTime();
}

function wmSortCurrentProgressRecentView_(rows) {
  /* WM_CURRENT_PROGRESS_RECENT_VIEW_ORDER_V1_20260821
   * 저장행은 학생ID 기준 1행 고정으로 유지하고 표시순서만 결정합니다.
   * 1차: 최종수정일 밀리초 내림차순 / 2차: 학생ID 오름차순.
   * 동일 시각 동시완료가 발생해도 항상 같은 순서를 재현합니다. */
  return (rows || []).slice().sort(function(a, b) {
    var at = wmCurrentProgressRecentSortTime_(a && a['최종수정일']);
    var bt = wmCurrentProgressRecentSortTime_(b && b['최종수정일']);
    if (at !== bt) return bt - at;
    var aid = String(a && a['학생ID'] || '').trim().toUpperCase();
    var bid = String(b && b['학생ID'] || '').trim().toUpperCase();
    return aid < bid ? -1 : (aid > bid ? 1 : 0);
  });
}

function getCurrentProgressForLmsApi_(e) {
  try {
    /* 고정행 DB의 물리적 행순서에 의존하지 않고 읽은 뒤 최신순 VIEW로 정렬합니다. */
    var result = wmReadLmsSheetAsObjects_('8.현재진행_DB');
    var actor = wmGetLmsRequestActor_(e);
    var rows = result.rows || [];
    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      rows = wmFilterRowsByStudentSet_(rows, wmAllowedStudentIdSetForActor_(actor) || {}, actor);
    }
    rows = wmSortCurrentProgressRecentView_(rows);
    if (wmIsLmsInitialPageRequest_(e)) {
      rows = rows.slice(0, wmLmsInitialLimit_(e, 25));
    }
    return outputResult(e, {
      success: !!result.success,
      message: result.message || '',
      currentProgress: rows
    });
  } catch (err) {
    return outputResult(e, { success:false, message:'8.현재진행_DB 조회 오류', error:String(err && err.message ? err.message : err), currentProgress:[] });
  }
}


/* WM_LMS_DASHBOARD_API_20260706_V1
 * LMS.html의 ensureDashboardSourceRows()가 기대하는 필드명 그대로 반환합니다.
 * 필수: students / records / currentProgress
 * 선택: classInfo / sendLog / counsel / payments
 */
function wmReadFirstExistingLmsSheetAsObjects_(sheetNames) {
  var names = sheetNames || [];
  for (var i = 0; i < names.length; i++) {
    var name = String(names[i] || '').trim();
    if (!name) continue;

    try {
      var result = wmReadLmsSheetAsObjects_(name);
      if (result && result.success) {
        return {
          success: true,
          sheetName: name,
          headers: result.headers || [],
          rows: result.rows || []
        };
      }
    } catch (err) {}
  }

  return {
    success: false,
    sheetName: '',
    headers: [],
    rows: []
  };
}

function wmNormalizeDashboardDateText_(value) {
  var text = String(value || '').trim();
  if (!text) return '';

  var date = null;

  try {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      date = value;
    }
  } catch (err1) {}

  if (!date) {
    var match = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
    if (match) {
      date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    }
  }

  if (!date) {
    var parsed = new Date(text);
    if (!isNaN(parsed.getTime())) date = parsed;
  }

  if (!date || isNaN(date.getTime())) return text;

  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
}

function wmDashboardIsToday_(value) {
  var rowDate = wmNormalizeDashboardDateText_(value);
  if (!rowDate) return false;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  return rowDate === today;
}

function wmDashboardCell_(row, names) {
  row = row || {};
  names = names || [];
  for (var i = 0; i < names.length; i++) {
    var key = String(names[i] || '').trim();
    if (key && row.hasOwnProperty(key) && String(row[key] || '').trim() !== '') return row[key];
  }
  return '';
}

function wmDashboardUniqueCount_(rows, names) {
  var map = {};
  var count = 0;
  rows = rows || [];
  for (var i = 0; i < rows.length; i++) {
    var value = String(wmDashboardCell_(rows[i], names) || '').trim();
    if (!value) continue;
    if (!map[value]) {
      map[value] = true;
      count++;
    }
  }
  return count;
}

function wmDashboardIsWithdrawStudent_(row) {
  var status = String(wmDashboardCell_(row, ['등록상태','활성상태','상태','운영상태','Status']) || '').trim().toLowerCase();
  return status.indexOf('퇴원') >= 0 ||
    status.indexOf('withdraw') >= 0 ||
    status.indexOf('leave') >= 0 ||
    status.indexOf('inactive') >= 0 ||
    status.indexOf('종료') >= 0;
}

function wmBuildDashboardSummary_(students, records, currentProgress, classInfo, teacherInfo, sendLog, counsel, payments) {
  students = students || [];
  records = records || [];
  currentProgress = currentProgress || [];
  classInfo = classInfo || [];
  teacherInfo = teacherInfo || [];
  sendLog = sendLog || [];
  counsel = counsel || [];
  payments = payments || [];

  var activeStudents = [];
  var activeStudentIdMap = {};
  for (var i = 0; i < students.length; i++) {
    if (wmDashboardIsWithdrawStudent_(students[i])) continue;
    activeStudents.push(students[i]);
    var sid = String(wmDashboardCell_(students[i], ['학생ID','studentId','Student_ID']) || '').trim();
    if (sid) activeStudentIdMap[sid] = true;
  }

  var todayStudyIdMap = {};
  var todayStudyRows = [];
  for (var r = 0; r < records.length; r++) {
    var studyDate = wmDashboardCell_(records[r], ['학습날짜','학습일','날짜','date','Date']);
    if (!wmDashboardIsToday_(studyDate)) continue;
    todayStudyRows.push(records[r]);
    var recordSid = String(wmDashboardCell_(records[r], ['학생ID','studentId','Student_ID']) || '').trim();
    if (recordSid) todayStudyIdMap[recordSid] = true;
  }

  var notStudyCount = 0;
  for (var s = 0; s < activeStudents.length; s++) {
    var activeSid = String(wmDashboardCell_(activeStudents[s], ['학생ID','studentId','Student_ID']) || '').trim();
    if (!activeSid || !todayStudyIdMap[activeSid]) notStudyCount++;
  }

  var classRowsForCount = classInfo.length ? classInfo : activeStudents;
  var teacherRowsForCount = teacherInfo.length ? teacherInfo : activeStudents;
  var leaderRowsForCount = teacherInfo.length ? teacherInfo : activeStudents;

  return {
    totalStudents: activeStudents.length,
    totalClasses: wmDashboardUniqueCount_(classRowsForCount, ['반ID','반명','현재반','Class','class','반']),
    totalTeachers: wmDashboardUniqueCount_(teacherRowsForCount, ['교사ID','이름','교사명','담당교사','현재담당교사','Teacher_Name']),
    totalLeaders: wmDashboardUniqueCount_(leaderRowsForCount, ['리더ID','담당리더','리더명','Leader_Name','이름']),
    todayStudy: Object.keys(todayStudyIdMap).length,
    todayStudyRows: todayStudyRows.length,
    notStudy: notStudyCount,
    currentProgressRows: currentProgress.length,
    sendLog: sendLog.length,
    counsel: counsel.length,
    payments: payments.length
  };
}


/* WM_LMS_DASHBOARD_PERSISTENT_SNAPSHOT_V2
 * 첫 화면은 ScriptProperties에 저장된 최소 스냅샷을 즉시 반환합니다.
 * 강제 새로고침(refresh=1)일 때만 원본 DB를 다시 읽어 스냅샷을 갱신합니다. */
function wmDashboardDataVersion_() {
  try {
    return String(
      PropertiesService.getScriptProperties()
        .getProperty('WM_DASHBOARD_DATA_VERSION_V1') || '0'
    );
  } catch (err) {
    return '0';
  }
}

function wmMarkDashboardDataChanged_() {
  var version = String(Date.now());
  try {
    PropertiesService.getScriptProperties()
      .setProperty('WM_DASHBOARD_DATA_VERSION_V1', version);
  } catch (err) {}
  return version;
}
function wmDashboardSnapshotPropertyKey_(actorKey, startText, endText) {
  var raw = [WM_ENV, actorKey || 'PUBLIC', startText || '', endText || '', wmDashboardDataVersion_()].join('|');
  var digest = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8)
  ).replace(/=+$/,'');
  return 'WM_DASHBOARD_SNAPSHOT_V2_' + digest;
}

function wmGetDashboardPersistentSnapshot_(key, maxAgeMs) {
  try {
    var text = PropertiesService.getScriptProperties().getProperty(key);
    if (!text) return null;
    var saved = JSON.parse(text);
    if (!saved || !saved.savedAt || !saved.payload || saved.payload.success !== true) return null;
    if (Date.now() - Number(saved.savedAt) > Number(maxAgeMs || 21600000)) return null;
    return saved.payload;
  } catch (err) {
    return null;
  }
}

function wmPutDashboardPersistentSnapshot_(key, payload) {
  try {
    var text = JSON.stringify({savedAt:Date.now(), payload:payload});
    if (text.length <= 480000) PropertiesService.getScriptProperties().setProperty(key, text);
  } catch (err) {}
}

function getDashboardForLmsApi_(e) {
  try {
    /* WM_LMS_DASHBOARD_MINIMUM_FIRST_V1
     * 최초 화면에서는 학생관리_DB와 학습기록_DB의 필요한 값만 읽습니다.
     * 현재진행·교사·반·발송·상담·결제 DB 전체 조회는 최초 진입에서 제외합니다. */
    var p = (e && e.parameter) ? e.parameter : {};
    var actor = wmGetLmsRequestActor_(e);
    var actorKey = actor && actor.found
      ? [wmNormalizeLmsRole_(actor.role), String(actor.teacherId || actor.id || '').trim().toUpperCase()].join(':')
      : 'PUBLIC';
    var startText = String(p.dateStart || '').trim();
    var endText = String(p.dateEnd || '').trim();
    var cacheKey = wmCacheKey_('LMS_DASHBOARD_MINIMUM', actorKey + ':' + startText + ':' + endText + ':' + wmDashboardDataVersion_());
    var propertyKey = wmDashboardSnapshotPropertyKey_(actorKey, startText, endText);
    var forceRefresh = String(p.refresh || '') === '1';
    if (!forceRefresh) {
      var cached = wmCacheGetJson_(cacheKey);
      if (cached && cached.success) {
        cached.snapshotSource = 'MEMORY_CACHE';
        return outputResult(e, cached);
      }
      var persistent = wmGetDashboardPersistentSnapshot_(propertyKey, 6 * 60 * 60 * 1000);
      if (persistent && persistent.success) {
        persistent.snapshotSource = 'SCRIPT_PROPERTIES';
        wmCachePutJson_(cacheKey, persistent, 600);
        return outputResult(e, persistent);
      }
    }

    var studentsResult = wmReadLmsSheetAsObjects_('1.학생관리_DB');
    var students = studentsResult.rows || [];
    if (actor && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role))) {
      students = wmFilterStudentsForActor_(students, actor);
    }

    function cell_(row, names) {
      return String(wmLmsCell_(row, names) || '').trim();
    }
    function dateKey_(value) {
      var text = String(value || '').trim();
      var match = text.match(/(\d{4})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})/);
      if (!match) return '';
      return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
    }
    function inPeriod_(value) {
      var key = dateKey_(value);
      if (!key) return false;
      if (startText && key < startText) return false;
      if (endText && key > endText) return false;
      return true;
    }
    function active_(row) {
      var status = cell_(row, ['활성상태','등록상태','상태','Status']).toUpperCase();
      return !status || status === 'ACTIVE' || status === '사용' || status === '재원' || status.indexOf('정상') >= 0;
    }

    var activeStudents = students.filter(active_);
    var activeIds = {};
    activeStudents.forEach(function(row){
      var id = cell_(row, ['학생ID','Student_ID','studentId']);
      if (id) activeIds[id.toUpperCase()] = true;
    });

    var recordSheet = getLmsSpreadsheet_().getSheetByName('2.학습기록_DB');
    var todayIds = {};
    var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (recordSheet && recordSheet.getLastRow() >= 2 && recordSheet.getLastColumn() >= 1) {
      var recordHeaders = recordSheet.getRange(1, 1, 1, recordSheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
      var idIndex = wmFindHeaderIndex_(recordHeaders, ['학생ID','Student_ID','studentId']);
      var dateIndex = wmFindHeaderIndex_(recordHeaders, ['학습날짜','학습일','날짜','date']);
      if (idIndex >= 0 && dateIndex >= 0) {
        var rowCount = recordSheet.getLastRow() - 1;
        var idValues = recordSheet.getRange(2, idIndex + 1, rowCount, 1).getDisplayValues();
        var dateValues = recordSheet.getRange(2, dateIndex + 1, rowCount, 1).getDisplayValues();
        for (var i = 0; i < rowCount; i++) {
          var sid = String(idValues[i][0] || '').trim();
          if (!sid || !activeIds[sid.toUpperCase()]) continue;
          if (dateKey_(dateValues[i][0]) === todayKey) todayIds[sid.toUpperCase()] = true;
        }
      }
    }

    var classSet = {};
    var teacherSet = {};
    var leaderSet = {};
    var monthlyNew = 0;
    var monthlyRegister = 0;
    var monthlyLeave = 0;
    activeStudents.forEach(function(row){
      var className = cell_(row, ['반명','현재반','Class','반']);
      var teacherName = cell_(row, ['교사명','현재담당교사','담당교사']);
      var leaderName = cell_(row, ['담당리더','리더','리더명']);
      if (className) classSet[className.toUpperCase()] = true;
      if (teacherName) teacherSet[teacherName.toUpperCase()] = true;
      if (leaderName) leaderSet[leaderName.toUpperCase()] = true;
      if (inPeriod_(cell_(row, ['등록일','최초등록일']))) monthlyNew++;
      if (inPeriod_(cell_(row, ['최근결제일','최종결제','등록일']))) monthlyRegister++;
    });
    students.forEach(function(row){
      var status = cell_(row, ['활성상태','등록상태','상태','Status']).toUpperCase();
      if ((status.indexOf('퇴원') >= 0 || status === 'INACTIVE') &&
          inPeriod_(cell_(row, ['퇴원일','종료일','수정일','등록일','최근결제일']))) monthlyLeave++;
    });

    var firstRows = activeStudents.slice().sort(function(a, b){
      return dateKey_(cell_(b, ['등록일','최초등록일'])) .localeCompare(dateKey_(cell_(a, ['등록일','최초등록일'])));
    }).slice(0, 15).map(function(row){
      return {
        '학생ID':cell_(row, ['학생ID','Student_ID','studentId']),
        '학생이름':cell_(row, ['학생이름','이름','Student_Name']),
        '반명':cell_(row, ['반명','현재반','Class','반']),
        '교사명':cell_(row, ['교사명','현재담당교사','담당교사']),
        '학교':cell_(row, ['학교']),
        '등록일':cell_(row, ['등록일','최초등록일']),
        '활성상태':cell_(row, ['활성상태','등록상태','상태'])
      };
    });

    var totalStudents = Object.keys(activeIds).length || activeStudents.length;
    var todayStudy = Object.keys(todayIds).length;
    var payload = {
      success:true,
      dashboardFast:true,
      message:'대시보드 최소 데이터 조회 성공',
      actor:{
        role:actor ? wmNormalizeLmsRole_(actor.role) : '',
        teacherId:actor ? String(actor.teacherId || '').trim() : '',
        name:actor ? String(actor.name || '').trim() : ''
      },
      summary:{
        totalStudents:totalStudents,
        totalClasses:Object.keys(classSet).length,
        totalTeachers:Object.keys(teacherSet).length,
        totalLeaders:Object.keys(leaderSet).length,
        monthlyNew:monthlyNew,
        monthlyRegister:monthlyRegister,
        monthlyLeave:monthlyLeave,
        todayStudy:todayStudy,
        notStudy:Math.max(0, totalStudents - todayStudy),
        todayRate:totalStudents ? Math.round(todayStudy * 100 / totalStudents) : 0
      },
      firstRows:firstRows
    };
    payload.snapshotSource = 'LIVE_DB';
    wmCachePutJson_(cacheKey, payload, 600);
    wmPutDashboardPersistentSnapshot_(propertyKey, payload);
    return outputResult(e, payload);
  } catch (err) {
    return outputResult(e, {
      success:false,
      dashboardFast:true,
      message:'대시보드 최소 데이터 조회 오류',
      error:String(err && err.message ? err.message : err),
      summary:{},
      firstRows:[]
    });
  }
}

function resolveClassDisplayName_(classId, className) {
  /* WM_CLASS_DISPLAY_NAME_ONLY_V1
   * 사용자 화면/학습기록/성적표에는 반ID가 아니라 실제 반명을 표시합니다.
   * 1.학생관리_DB에 반명 계열 컬럼이 있으면 그 값을 우선 사용하고,
   * 없으면 7-1.반관리_DB의 반ID→반명 매핑으로 보강합니다.
   */
  var displayName = String(className || '').trim();
  if (displayName) {
    return displayName;
  }

  var internalId = String(classId || '').trim();
  if (!internalId) {
    return '';
  }

  var mappedName = lookupClassNameById_(internalId);
  if (mappedName) {
    return mappedName;
  }

  return internalId;
}

function lookupClassNameById_(classId) {
  var wantedId = String(classId || '').trim();
  if (!wantedId) {
    return '';
  }

  try {
    var ss = getLmsSpreadsheet_();
    if (!ss) {
      return '';
    }

    var sheet = ss.getSheetByName('7-2.반관리_DB');
    if (!sheet) {
      return '';
    }

    var values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 2) {
      return '';
    }

    var headers = values[0].map(function(h){ return String(h || '').trim(); });
    var idxClassId = findFirstHeaderIndex_(headers, ['반ID', 'Class', '클래스ID']);
    var idxClassName = findFirstHeaderIndex_(headers, ['반명', '학급명', '클래스명', '수업반명', '반이름']);

    if (idxClassId < 0 || idxClassName < 0) {
      return '';
    }

    for (var i = 1; i < values.length; i++) {
      var row = values[i] || [];
      var rowClassId = String(row[idxClassId] || '').trim();
      if (rowClassId === wantedId) {
        return String(row[idxClassName] || '').trim();
      }
    }
  } catch (err) {}

  return '';
}


function getStudentMapProfile(studentId) {
  var info = getStudentBasicInfoForMap_(studentId);
  var cache = wmGetCurrentProgressCacheForStudent_(studentId);
  var profile = wmApplyCurrentProgressCacheToStudentProfile_({
    학생ID: info.studentId || String(studentId || '').trim(),
    학생이름: info.studentName || '',
    학교: info.school || '',
    학년: info.grade || '',
    Class: info.className || '',
    교사명: info.teacherName || '',
    학습배정: info.learningAssign || '',
    현재세트: info.currentSet || '',
    학습모드: info.learningMode || buildDefaultLearningMode_()
  }, cache);

  return {
    success: !!profile.학생ID,
    학생ID: profile.학생ID || '',
    학생이름: profile.학생이름 || '',
    학교: profile.학교 || '',
    학년: profile.학년 || '',
    Class: profile.Class || '',
    교사명: profile.교사명 || '',
    학습배정: profile.학습배정 || '',
    현재세트: profile.현재세트 || '',
    학습모드: profile.학습모드 || buildDefaultLearningMode_(),
    레벨완료조건: profile.레벨완료조건 || '',
    레벨회차추가: profile.레벨회차추가 || '',
    레벨완료횟수: profile.레벨완료횟수 || 0,
    순차완주세트: profile.순차완주세트 || 0,
    순차완주회차: profile.순차완주회차 || 0
  };
}

/* WM_STUDENT_ID_RULE_260708_V1
 * 학생ID 공식: 영문/숫자 4~10자, 입력값은 그대로 유지하고 비교만 대소문자를 구분하지 않습니다.
 * 기존 S4821/S1057 형식은 그대로 유효합니다.
 */
function wmNormalizeStudentId_(value) {
  return String(value || '').trim();
}

function wmIsValidStudentId_(value) {
  return /^[A-Za-z0-9]{4,10}$/.test(wmNormalizeStudentId_(value));
}

function wmStudentIdCompareKey_(value) {
  return wmNormalizeStudentId_(value).toLowerCase();
}

/* WM_STUDENT_ROW_INDEX_CACHE_SPEED_V1_20260826
 * 학생ID 1열만 읽어 ID→행번호 인덱스를 공용 캐시합니다.
 * 특정 학생ID 예외 없이 로그인/현재진행 조회가 같은 행찾기 공식을 사용합니다.
 * 캐시 miss 또는 새 학생으로 ID가 없을 때만 학생ID 1열을 다시 읽습니다. */
function wmGetStudentRowNumberByIdIndex_(sheet, idColumn, lastRow, cacheName, studentId, forceRefresh) {
  if (!sheet || !idColumn || lastRow < 2) return 0;
  var wanted = wmStudentIdCompareKey_(studentId);
  if (!wanted) return 0;

  var cacheKey = wmCacheKey_(String(cacheName || 'STUDENT_ROW_INDEX'), 'ALL');
  var cached = forceRefresh ? null : wmCacheGetJson_(cacheKey);
  var rows = cached && cached.rows && Number(cached.lastRow || 0) === Number(lastRow)
    ? cached.rows
    : null;

  if (!rows || !Object.prototype.hasOwnProperty.call(rows, wanted)) {
    var values = sheet.getRange(2, idColumn, lastRow - 1, 1).getDisplayValues();
    rows = {};
    for (var i = 0; i < values.length; i++) {
      var key = wmStudentIdCompareKey_(values[i][0]);
      if (key && !Object.prototype.hasOwnProperty.call(rows, key)) {
        rows[key] = i + 2;
      }
    }
    wmCachePutJson_(cacheKey, { lastRow:lastRow, rows:rows }, 1800);
  }

  return Number(rows[wanted] || 0);
}

function wmGenerateStudentIdForLms_() {
  var ss = getLmsSpreadsheet_();
  if (!ss) return { success:false, message:'스프레드시트 연결 실패' };

  var sheet = ss.getSheetByName('1.학생관리_DB');
  if (!sheet) return { success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { success:false, message:'학생ID 컬럼을 확인할 수 없습니다.' };

  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var idxStudentId = headers.indexOf('학생ID');
  if (idxStudentId === -1) return { success:false, message:'학생ID 컬럼을 찾을 수 없습니다.' };

  var used = {};
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, idxStudentId + 1, lastRow - 1, 1).getDisplayValues();
    ids.forEach(function(row) {
      var id = wmStudentIdCompareKey_(row[0]);
      if (id) used[id] = true;
    });
  }

  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var i = 0; i < 5000; i++) {
    var id = '';
    for (var j = 0; j < 6; j++) {
      id += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    if (wmIsValidStudentId_(id) && !used[wmStudentIdCompareKey_(id)]) {
      return { success:true, studentId:id, rule:'영문/숫자 4~10자, 대소문자 구분 없음, 중복 확인' };
    }
  }

  return { success:false, message:'사용 가능한 학생ID를 생성하지 못했습니다.' };
}

function studentLogin(e) {
  /* WM_STUDENT_LOGIN_5COLUMN_DIRECT_V1_20260828
   * B 학생ID 검색 → D 비밀번호 + AE ACTIVE 확인 → AH 현재세션 + AI LOGIN 저장 → 즉시 성공.
   * 캐시/학생정보/전체행조회/기존세션확인/중복로그인확인은 사용하지 않습니다.
   */
  var studentId = '';
  var password = '';
  var deviceInfo = '';

  if (e && e.parameter) {
    studentId = String(e.parameter.studentId || '').trim();
    password = String(e.parameter.password || '').trim();
    deviceInfo = String(e.parameter.deviceInfo || e.parameter.userAgent || '').trim();
  }

  if (!studentId || !password) {
    return outputResult(e, { success:false, message:'학생ID와 비밀번호를 입력하세요.' });
  }

  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('1.학생관리_DB');
  if (!sheet) {
    return outputResult(e, { success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.' });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return outputResult(e, { success:false, message:'학생 데이터가 없습니다.' });
  }

  var finder = sheet
    .getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(studentId)
    .matchEntireCell(true);
  try { finder.matchCase(false); } catch (ignoreMatchCase) {}

  var idCell = finder.findNext();
  if (!idCell) {
    return outputResult(e, { success:false, message:'학생ID 또는 비밀번호가 일치하지 않습니다.' });
  }

  var rowNumber = idCell.getRow();
  var loginRow = sheet.getRange(rowNumber, 2, 1, 3).getDisplayValues()[0], dbStudentId = String(loginRow[0] || '').trim(), rowPassword = String(loginRow[2] || '').trim();
  var activeDeviceRange = sheet.getRange(rowNumber, 31, 1, 7).getDisplayValues()[0];
  var activeStatus = String(activeDeviceRange[0] || '').trim().toUpperCase();
  var currentDevice = String(activeDeviceRange[6] || '').trim();

  if (rowPassword !== password) {
    return outputResult(e, { success:false, message:'학생ID 또는 비밀번호가 일치하지 않습니다.' });
  }

  if (activeStatus === 'INACTIVE') {
    return outputResult(e, { success:false, message:'현재 로그인 가능한 학생 상태가 아닙니다.' });
  }

  var sessionToken = Utilities.getUuid();
  var normalizedDevice = normalizeDeviceInfo_(deviceInfo) || String(deviceInfo || '').trim() || 'PC';
  sheet.getRange(rowNumber, 34, 1, 3).setValues([[sessionToken, 'LOGIN', Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')]]);
  if (currentDevice !== normalizedDevice) {
    sheet.getRange(rowNumber, 37).setValue(normalizedDevice);
  }

  return outputResult(e, {
    success:true,
    message:'로그인 성공',
    sessionToken:sessionToken,
    student:{
      학생ID:dbStudentId,
      현재세션:sessionToken,
      sessionToken:sessionToken
    }
  });
}
function checkStudentSession(e) {
  try {
    var studentId = '';
    var sessionToken = '';

    if (e && e.parameter) {
      studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || '').trim();
      sessionToken = String(e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN || '').trim();
    }

    var state = getStudentSessionState_(studentId);

    if (!studentId || !sessionToken) {
      return outputResult(e, {
        success: false,
        sessionValid: false,
        message: '세션 정보가 없습니다.'
      });
    }

    if (!state.success) {
      return outputResult(e, state);
    }

    var valid = !!state.currentSession && state.currentSession === sessionToken && String(state.sessionStatus || '').toUpperCase() === 'LOGIN';

    return outputResult(e, {
      success: true,
      sessionValid: valid,
      valid: valid,
      message: valid ? '세션 정상' : '다른 기기에서 로그인되었습니다.',
      studentId: studentId
    });
  } catch (err) {
    return outputResult(e, {
      success: false,
      sessionValid: false,
      message: '세션 확인 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}

function studentLogout(e) {
  try {
    var studentId = '';
    var sessionToken = '';

    if (e && e.parameter) {
      studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || '').trim();
      sessionToken = String(e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN || '').trim();
    }

    var state = getStudentSessionState_(studentId);
    if (!state.success) {
      return outputResult(e, state);
    }

    if (sessionToken && state.currentSession && state.currentSession !== sessionToken) {
      return outputResult(e, {
        success: true,
        message: '이미 다른 세션으로 변경되었습니다.',
        cleared: false
      });
    }

    state.sheet.getRange(state.rowNumber, state.idxCurrentSession + 1, 1, 2)
      .setValues([['', 'LOGOUT']]);

    return outputResult(e, {
      success: true,
      message: '로그아웃 처리 완료',
      cleared: true
    });
  } catch (err) {
    return outputResult(e, {
      success: false,
      message: '로그아웃 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}

function getStudentSessionState_(studentId) {
  /* WM_STUDENT_SESSION_FIXED_COL_FAST_V2_20260818
   * 1.학생관리_DB 공식 고정열(B 학생ID / AH 현재세션 / AI 세션상태)을 사용합니다.
   * 로그인에서 만든 학생행 캐시를 재사용하고, 전체 헤더/전체 학생행 조회를 하지 않습니다.
   */
  studentId = String(studentId || '').trim();

  if (!studentId) {
    return {
      success: false,
      message: 'studentId가 없습니다.'
    };
  }

  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('1.학생관리_DB');

  if (!sheet) {
    return {
      success: false,
      message: '1.학생관리_DB 시트를 찾을 수 없습니다.'
    };
  }

  var COL_STUDENT_ID = 2;
  var COL_CURRENT_SESSION = 34;
  var COL_SESSION_STATUS = 35;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: false,
      message: '학생 정보가 없습니다.'
    };
  }

  var rowCacheKey = wmCacheKey_('LOGIN_ROW_5FIELD', wmStudentIdCompareKey_(studentId));
  var cachedRowNumber = Number(wmCacheGetJson_(rowCacheKey) || 0);
  var rowNumber = 0;
  var cachedRowValues = null;

  if (cachedRowNumber >= 2 && cachedRowNumber <= lastRow) {
    cachedRowValues = sheet.getRange(cachedRowNumber, COL_STUDENT_ID, 1, COL_SESSION_STATUS - COL_STUDENT_ID + 1).getDisplayValues()[0];
    var cachedId = String(cachedRowValues[0] || '').trim();
    if (wmStudentIdCompareKey_(cachedId) === wmStudentIdCompareKey_(studentId)) {
      rowNumber = cachedRowNumber;
    } else {
      cachedRowValues = null;
    }
  }

  if (!rowNumber) {
    var finder = sheet
      .getRange(2, COL_STUDENT_ID, lastRow - 1, 1)
      .createTextFinder(studentId)
      .matchEntireCell(true);
    try { finder.matchCase(false); } catch (caseErr) {}
    var cell = finder.findNext();
    if (!cell) {
      return {
        success: false,
        message: '학생 정보를 찾을 수 없습니다.'
      };
    }
    rowNumber = cell.getRow();
    wmCachePutJson_(rowCacheKey, rowNumber, 1800);
  }

  var currentSession = '';
  var sessionStatus = '';
  if (cachedRowValues) {
    currentSession = String(cachedRowValues[COL_CURRENT_SESSION - COL_STUDENT_ID] || '').trim();
    sessionStatus = String(cachedRowValues[COL_SESSION_STATUS - COL_STUDENT_ID] || '').trim();
  } else {
    var stateValues = sheet.getRange(rowNumber, COL_CURRENT_SESSION, 1, 2).getDisplayValues()[0];
    currentSession = String(stateValues[0] || '').trim();
    sessionStatus = String(stateValues[1] || '').trim();
  }

  return {
    success: true,
    sheet: sheet,
    rowNumber: rowNumber,
    idxCurrentSession: COL_CURRENT_SESSION - 1,
    idxSessionStatus: COL_SESSION_STATUS - 1,
    currentSession: currentSession,
    sessionStatus: sessionStatus
  };
}

function wmBuildSessionExpiredResponse_() {
  return {
    success: false,
    sessionValid: false,
    valid: false,
    sessionExpired: true,
    message: '다른 기기에서 로그인되었습니다.'
  };
}

function wmIsStudentSessionValid_(studentId, sessionToken) {
  studentId = String(studentId || '').trim().toUpperCase();
  sessionToken = String(sessionToken || '').trim();

  if (!studentId || !sessionToken) {
    return true;
  }

  var state = getStudentSessionState_(studentId);
  if (!state || !state.success) {
    return false;
  }

  return !!state.currentSession
    && state.currentSession === sessionToken
    && String(state.sessionStatus || '').toUpperCase() === 'LOGIN';
}

function wmAssertStudentSessionFromParams_(e) {
  var p = e && e.parameter ? e.parameter : {};
  var studentId = String(p.studentId || p.studentID || p.sid || '').trim().toUpperCase();
  var sessionToken = String(p.sessionToken || p.wmSessionToken || p.WM_SESSION_TOKEN || '').trim();

  if (!studentId && !sessionToken) {
    return { ok: true };
  }

  if (!studentId || !sessionToken) {
    return { ok: false, response: wmBuildSessionExpiredResponse_() };
  }

  if (!wmIsStudentSessionValid_(studentId, sessionToken)) {
    return { ok: false, response: wmBuildSessionExpiredResponse_() };
  }

  return { ok: true };
}

function buildInitialSetIdFromLearningAssign(learningText) {
  var text = String(learningText || '').trim();
  var match = text.match(/(\d+)/);

  if (!match) {
    return '5-1-1';
  }

  return Number(match[1]) + '-1-1';
}


function normalizeStudySetId(setId) {
  var text = String(setId || '').trim().toUpperCase();

  if (/^WM\d+-\d+-\d+$/.test(text)) {
    return text;
  }

  if (/^\d+-\d+-\d+$/.test(text)) {
    return 'WM' + text;
  }

  return 'WM5-1-1';
}

function startLearning(e) {
  try {
    var studentId = '';
    var setId = '';
    var studentName = '';

    if (e && e.parameter) {
      studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || '').trim();
      setId = String(e.parameter.setId || e.parameter.set_id || e.parameter.Set_ID || '').trim().toUpperCase();
      studentName = String(e.parameter.studentName || e.parameter.name || '').trim();
      var sessionToken = String(e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN || '').trim();
    }

    if (!studentId) {
      return outputResult(e, {
        success: false,
        message: 'studentId가 없습니다.'
      });
    }

    if (sessionToken && !wmIsStudentSessionValid_(studentId, sessionToken)) {
      return outputResult(e, wmBuildSessionExpiredResponse_());
    }

    setId = normalizeStudySetId(setId);

    var baseUrl = ScriptApp.getService().getUrl();
    var studyUrl = baseUrl
      + '?mode=study'
      + '&studentId=' + encodeURIComponent(studentId)
      + '&setId=' + encodeURIComponent(setId)
      + '&studentName=' + encodeURIComponent(studentName)
      + '&sessionToken=' + encodeURIComponent(sessionToken || '');

    return outputResult(e, {
      success: true,
      message: '학습 시작 준비 완료',
      studentId: studentId,
      setId: setId,
      studyUrl: studyUrl
    });

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: '학습 시작 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}


function wmBuildFastLearningMapPayloadFromCurrentProgress_(studentId, studentProfile, currentProgressCache, currentProgressState) {
  currentProgressCache = currentProgressCache || null;
  if (!currentProgressCache) return null;

  studentProfile = studentProfile || buildStudentProfileForLearningMap_(studentId);
  currentProgressState = currentProgressState || wmBuildCurrentProgressLearningStateForMap_(studentProfile, currentProgressCache);
  var sourceSetId = wmNormalizeCurrentProgressRecordSetId_(currentProgressCache['Set_ID'] || '', '');
  var sourceCurrentStep = String(currentProgressCache['현재Step'] || '').trim().toUpperCase();
  if (!sourceSetId || !sourceCurrentStep || !currentProgressState || !currentProgressState.currentSet || !currentProgressState.officialLevelProgress || !currentProgressState.officialLevelProgress.level) return null;

  var recentText = String(currentProgressCache['최근30건JSON'] || '').trim();
  var recentList = [];
  if (recentText) {
    try { recentList = JSON.parse(recentText); } catch (err) { recentList = []; }
  }
  if (!Array.isArray(recentList)) recentList = [];

  var records = [];
  var completedSetIds = [];
  var setStatusMap = {};

  recentList.forEach(function(item) {
    item = item || {};
    var setId = normalizeLevelPlainSetId_(item.Set_ID || item['Set_ID'] || item.setId || '');
    if (!setId) return;

    var status = String(item['완료상태'] || item.status || '').trim();
    var record = {
      학습날짜: item['학습날짜'] || '',
      학습기록ID: item['학습기록ID'] || '',
      학생ID: studentId,
      학생이름: studentProfile['학생이름'] || '',
      학교: studentProfile['학교'] || '',
      학년: studentProfile['학년'] || '',
      Class: studentProfile['Class'] || '',
      교사명: studentProfile['교사명'] || '',
      Set_ID: setId,
      완료상태: status,
      완료: item['완료'] || '',
      점수: item['점수'] || '',
      한영주관식: item['점수'] || ''
    };
    records.push(record);

    if (!setStatusMap[setId]) {
      setStatusMap[setId] = {
        setId: setId,
        repeatCount: 0,
        officialRound: 0,
        latestStatus: '',
        latestScore: '',
        latestDate: '',
        currentProgressStep: 'STEP1',
        lastCompletedStep: '',
        completedStepCount: 0,
        progressPercent: 0,
        progressRank: 0
      };
    }

    if (wmSafeIsCompleteStatusForCurrentProgress_(status)) {
      setStatusMap[setId].repeatCount += 1;
      setStatusMap[setId].officialRound = Math.max(Number(setStatusMap[setId].officialRound || 0), Number(String(item['완료'] || '').replace(/[^0-9]/g, '')) || 0);
      setStatusMap[setId].latestStatus = '완료';
      setStatusMap[setId].latestScore = item['점수'] || '';
      setStatusMap[setId].latestDate = item['학습날짜'] || '';
      setStatusMap[setId].currentProgressStep = 'COMPLETE';
      setStatusMap[setId].lastCompletedStep = 'TEST';
      setStatusMap[setId].completedStepCount = 6;
      setStatusMap[setId].progressPercent = 100;
      setStatusMap[setId].progressRank = 6;
      if (completedSetIds.indexOf(setId) === -1) completedSetIds.push(setId);
    }
  });

  Object.keys(currentProgressState.setStatusMap || {}).forEach(function(key) {
    setStatusMap[key] = currentProgressState.setStatusMap[key];
  });

  var officialLevelProgress = currentProgressState.officialLevelProgress || {};
  officialLevelProgress = wmApplyCurrentProgressCacheToOfficialLevelProgress_(officialLevelProgress, currentProgressCache);
  var currentLearningProgress = currentProgressState.currentLearningProgress || {};

  var levelHistoryLevels = [];
  try {
    levelHistoryLevels = buildLevelHistoryLevelsForLearningMap_(records, completedSetIds, setStatusMap, studentProfile, officialLevelProgress);
  } catch (err) {
    levelHistoryLevels = [];
  }

  var currentProgressHistoryCache = {
    levels: (String(currentProgressCache['히스토리레벨'] || currentProgressCache['완료된레벨'] || '').trim() || '').split('|').map(function(level){ return Number(String(level || '').replace(/[^0-9]/g, '')); }).filter(function(level){ return level >= 3 && level <= 13; }),
    roundsByLevel: {}
  };
  try {
    var __levels = currentProgressHistoryCache.levels;
    var __rounds = (String(currentProgressCache['히스토리횟수'] || currentProgressCache['완료횟수'] || '').trim() || '').split('|');
    for (var __hi = 0; __hi < __levels.length; __hi++) {
      currentProgressHistoryCache.roundsByLevel[__levels[__hi]] = Number(String(__rounds[__hi] || '').replace(/[^0-9]/g, '') || 0);
    }
  } catch (__historyCacheErr) {}

  return {
    success: true,
    message: '학습맵 빠른 조회 성공',
    studentId: studentId,
    studentProfile: studentProfile,
    currentProgressCache: currentProgressCache,
    currentProgress: currentProgressCache,
    currentProgressRow: currentProgressCache,
    progressCache: currentProgressCache,
    mapCache: currentProgressCache,
    records: records,
    completedSetIds: completedSetIds,
    officialLevelProgress: officialLevelProgress,
    currentLearningProgress: currentLearningProgress,
    levelHistoryLevels: levelHistoryLevels,
    currentProgressHistoryCache: currentProgressHistoryCache,
    setStatusMap: setStatusMap,
    fastMap: true
  };
}


/* WM_MAP_LEARNING_RECORD_DB_DIRECT_ATTACH_20260704_V1
 * 학습맵 하단 학습기록 영역은 8.현재진행_DB 최근30건JSON이 비어 있어도
 * 실제 원본인 2.학습기록_DB를 학생ID 기준으로 직접 읽어 payload.records에 채웁니다.
 * 상단/진행률 캐시는 기존 8.현재진행_DB 흐름을 유지하고, 하단 기록 목록만 보강합니다.
 */
function wmAttachLearningRecordDbRecordsForMap_(payload, studentId, studentProfile) {
  payload = payload || {};
  studentId = String(studentId || '').trim();
  studentProfile = studentProfile || {};

  if (!studentId) {
    return payload;
  }

  var mapRecordsCacheKey = wmCacheKey_('MAP_RECORDS', studentId);
  var cachedMapRecords = wmCacheGetJson_(mapRecordsCacheKey);
  if (cachedMapRecords && Array.isArray(cachedMapRecords.records)) {
    payload.records = cachedMapRecords.records;
    payload.learningRecords = cachedMapRecords.records;
    payload.recordsAll = cachedMapRecords.records;
    payload.learningRecordDbDirect = true;
    payload.learningRecordDbDirectCount = cachedMapRecords.records.length;
    payload.learningRecordDbCache = true;
    return payload;
  }

  try {
    var ss = getLmsSpreadsheet_();
    if (!ss) {
      payload.learningRecordDbDirect = false;
      payload.learningRecordDbDirectMessage = '스프레드시트 연결 실패';
      return payload;
    }

    var sheet = ss.getSheetByName('2.학습기록_DB');
    if (!sheet) {
      payload.learningRecordDbDirect = false;
      payload.learningRecordDbDirectMessage = '2.학습기록_DB 시트를 찾을 수 없습니다.';
      return payload;
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1 || sheet.getLastRow() < 2) {
      payload.records = [];
      payload.learningRecordDbDirect = true;
      payload.learningRecordDbDirectCount = 0;
      return payload;
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });

    var idxStudentId = headers.indexOf('학생ID');
    var idxSetId = headers.indexOf('Set_ID');

    if (idxStudentId < 0) {
      payload.learningRecordDbDirect = false;
      payload.learningRecordDbDirectMessage = '2.학습기록_DB에서 학생ID 컬럼을 찾을 수 없습니다.';
      return payload;
    }

    if (idxSetId < 0) {
      payload.learningRecordDbDirect = false;
      payload.learningRecordDbDirectMessage = '2.학습기록_DB에서 Set_ID 컬럼을 찾을 수 없습니다.';
      return payload;
    }

    var studentRowNumbers = wmGetStudentRowNumbersFast_(sheet, headers, studentId);
    var studentRows = wmReadGroupedRowsFast_(sheet, studentRowNumbers, lastCol);
    var records = [];

    for (var i = 0; i < studentRows.length; i++) {
      var row = studentRows[i] || [];
      var rowStudentId = String(row[idxStudentId] || '').trim();
      if (wmStudentIdCompareKey_(rowStudentId) !== wmStudentIdCompareKey_(studentId)) {
        continue;
      }

      var rawRecord = buildRecordObjectByHeaders_(headers, row);
      rawRecord.학생ID = rowStudentId;
      rawRecord.Set_ID = String(rawRecord.Set_ID || row[idxSetId] || '').trim();

      if (!rawRecord.Set_ID) {
        continue;
      }

      /* DB 원본을 수정하지 않고 화면 표시용 payload에만 학생 기본정보를 보강합니다. */
      rawRecord.학생이름 = String(rawRecord.학생이름 || studentProfile.학생이름 || studentProfile.studentName || '').trim();
      rawRecord.학교 = String(rawRecord.학교 || studentProfile.학교 || studentProfile.school || '').trim();
      rawRecord.학년 = String(rawRecord.학년 || studentProfile.학년 || studentProfile.grade || '').trim();
      rawRecord.Class = String(rawRecord.Class || rawRecord['반'] || rawRecord['반명'] || studentProfile.Class || studentProfile.className || '').trim();
      rawRecord.교사명 = String(rawRecord.교사명 || studentProfile.교사명 || studentProfile.teacherName || '').trim();
      rawRecord.__displayRecordId = String(rawRecord['학습기록ID'] || '').trim();

      records.push(buildLearningMapDisplayRecord_(rawRecord));
    }

    records.sort(function(a, b) {
      var at = getLearningRecordSortTime_(a && a['학습날짜']);
      var bt = getLearningRecordSortTime_(b && b['학습날짜']);
      return bt - at;
    });

    payload.records = records;
    payload.learningRecords = records;
    payload.recordsAll = records;
    payload.learningRecordDbDirect = true;
    payload.learningRecordDbDirectCount = records.length;
    payload.learningRecordDbCache = false;
    wmCachePutJson_(mapRecordsCacheKey, { records: records }, 60);
  } catch (err) {
    payload.learningRecordDbDirect = false;
    payload.learningRecordDbDirectError = String(err && err.message ? err.message : err);
  }

  return payload;
}

/* WM_MAP_RECORD_FIRST_PAGE_ONLY_V1_20260822
 * 최초 Map 진입에서는 학습기록 전체행을 읽지 않고 요청한 페이지(기본 15행)만 읽습니다.
 * 학생ID/학습날짜 인덱스 열만 확인한 뒤 해당 페이지의 실제 행만 상세 조회합니다.
 * pageSize가 없으면 기존 전체조회 공식을 그대로 사용합니다. */
function wmMapRecordDateText_(value) {
  var time = getLearningRecordSortTime_(value);
  if (!time) return '';
  return Utilities.formatDate(
    new Date(time),
    Session.getScriptTimeZone() || 'Asia/Seoul',
    'yyyy-MM-dd'
  );
}

function wmResolveMapRecordOneMonthRange_(fromDate, toDate) {
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  var todayText = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var rawFromText = String(fromDate || '').trim();
  var rawToText = String(toDate || '').trim();

  if (!rawFromText && !rawToText) {
    return { ok:true, apply:false, fromDate:'', toDate:'' };
  }

  var endText = rawToText || todayText;
  var startText = rawFromText;

  function parseIso_(text) {
    var m = String(text || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (d.getFullYear() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) return null;
    return d;
  }

  var endDate = parseIso_(endText);
  if (!endDate) return { ok:false, message:'학습기록 검색 종료일 형식이 올바르지 않습니다.' };

  if (!startText) {
    var defaultStart = new Date(endDate.getTime());
    defaultStart.setMonth(defaultStart.getMonth() - 1);
    startText = Utilities.formatDate(defaultStart, tz, 'yyyy-MM-dd');
  }

  var startDate = parseIso_(startText);
  if (!startDate) return { ok:false, message:'학습기록 검색 시작일 형식이 올바르지 않습니다.' };
  if (endDate.getTime() < startDate.getTime()) {
    return { ok:false, message:'끝일은 시작일보다 빠를 수 없습니다.' };
  }

  var oneMonthLimit = new Date(startDate.getTime());
  oneMonthLimit.setMonth(oneMonthLimit.getMonth() + 1);
  if (endDate.getTime() > oneMonthLimit.getTime()) {
    return { ok:false, message:'학습기록 검색은 최대 1개월 단위로 가능합니다.' };
  }

  return {
    ok:true,
    apply:true,
    fromDate:Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    toDate:Utilities.formatDate(endDate, tz, 'yyyy-MM-dd')
  };
}

/* WM_MAP_RECORD_1MONTH_HASMORE_ONLY_V1_20260822
 * 전체 학생행/전체 페이지수를 먼저 계산하지 않습니다.
 * 최근순으로 필요한 행만 200행 단위로 확인하고, 요청 페이지 15개 + 다음 1개를 찾는 즉시 중단합니다.
 * 검색 범위는 서버에서도 최대 1개월로 강제하며 응답은 hasMore만 반환합니다. */
function wmAttachLearningRecordDbPageForMap_(payload, studentId, studentProfile, page, pageSize, fromDate, toDate, searchFilterJson) {
  payload = payload || {};
  studentId = String(studentId || '').trim();
  studentProfile = studentProfile || {};
  page = Math.max(1, Math.floor(Number(page || 1) || 1));
  pageSize = Math.max(1, Math.min(100, Math.floor(Number(pageSize || 15) || 15)));

  if (!studentId) return payload;

  var dateRange = wmResolveMapRecordOneMonthRange_(fromDate, toDate);
  if (!dateRange.ok) {
    payload.records = [];
    payload.learningRecords = [];
    payload.recordsAll = [];
    payload.learningRecordDbDirect = false;
    payload.learningRecordDbPaged = true;
    payload.learningRecordDbPage = page;
    payload.learningRecordDbPageSize = pageSize;
    payload.learningRecordDbHasMore = false;
    payload.learningRecordDbRangeError = true;
    payload.message = dateRange.message;
    return payload;
  }

  var searchFilter = {};
  try {
    searchFilter = searchFilterJson
      ? (typeof searchFilterJson === 'string' ? JSON.parse(searchFilterJson) : searchFilterJson)
      : {};
  } catch (filterErr) {
    searchFilter = {};
  }

  var wantedSet = String(searchFilter.setKeyword || '').trim().toUpperCase().replace(/^WM/, '');
  var wantedLevel = String(searchFilter.levelKeyword || '').trim();
  var rawWantedMonth = String(searchFilter.month || '').trim();
  var wantedMonth = rawWantedMonth ? rawWantedMonth.padStart(2, '0') : '';
  var wantedWeek = String(searchFilter.week || '').trim();
  var needTotalCount = searchFilter.needTotalCount === true || String(searchFilter.needTotalCount || '').toLowerCase() === 'true';

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss && ss.getSheetByName('2.학습기록_DB');
    if (!sheet || sheet.getLastColumn() < 1 || sheet.getLastRow() < 2) {
      payload.records = [];
      payload.learningRecords = [];
      payload.recordsAll = [];
      payload.learningRecordDbDirect = true;
      payload.learningRecordDbPaged = true;
      payload.learningRecordDbPage = page;
      payload.learningRecordDbPageSize = pageSize;
      payload.learningRecordDbHasMore = false;
      return payload;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });
    var idxStudentId = headers.indexOf('학생ID');
    var idxSetId = headers.indexOf('Set_ID');
    var idxDate = headers.indexOf('학습날짜');
    if (idxStudentId < 0 || idxSetId < 0 || idxDate < 0) {
      payload.learningRecordDbDirect = false;
      payload.learningRecordDbDirectMessage = '2.학습기록_DB 필수 컬럼을 찾을 수 없습니다.';
      return payload;
    }

    var skipCount = (page - 1) * pageSize;
    var neededCount = skipCount + pageSize + 1;
    var candidateRows = [];
    var totalMatchedCount = 0;
    var scanRow = 2;
    var chunkSize = 200;
    var minIndex = Math.min(idxStudentId, idxSetId, idxDate);
    var maxIndex = Math.max(idxStudentId, idxSetId, idxDate);
    var scanWidth = maxIndex - minIndex + 1;
    var studentCompareKey = wmStudentIdCompareKey_(studentId);

    while (scanRow <= lastRow && (needTotalCount || candidateRows.length < neededCount)) {
      var scanCount = Math.min(chunkSize, lastRow - scanRow + 1);
      var scanValues = sheet.getRange(scanRow, minIndex + 1, scanCount, scanWidth).getDisplayValues();

      for (var s = 0; s < scanValues.length; s++) {
        var scan = scanValues[s] || [];
        var rowStudentId = String(scan[idxStudentId - minIndex] || '').trim();
        if (wmStudentIdCompareKey_(rowStudentId) !== studentCompareKey) continue;

        var dateText = wmMapRecordDateText_(scan[idxDate - minIndex]);
        if (dateRange.apply && (!dateText || dateText < dateRange.fromDate || dateText > dateRange.toDate)) continue;

        var setId = String(scan[idxSetId - minIndex] || '').trim().toUpperCase().replace(/^WM/, '');
        if (!setId) continue;
        if (wantedSet && setId !== wantedSet) continue;

        var levelMatch = setId.match(/^(3|4|5|6|7|8|9|10|11|12|13)-\d+-\d+$/);
        var rowLevel = levelMatch ? levelMatch[1] : '';
        if (wantedLevel && rowLevel !== wantedLevel) continue;

        if (wantedMonth && (!dateText || dateText.slice(5, 7) !== wantedMonth)) continue;
        if (wantedWeek) {
          if (!dateText) continue;
          var day = Number(dateText.slice(8, 10) || 0);
          var rowWeek = String(Math.min(5, Math.ceil(day / 7)));
          if (rowWeek !== wantedWeek) continue;
        }

        totalMatchedCount++;
        if (candidateRows.length < neededCount) candidateRows.push(scanRow + s);
        if (!needTotalCount && candidateRows.length >= neededCount) break;
      }

      scanRow += scanCount;
    }

    var pageRowNumbers = candidateRows.slice(skipCount, skipCount + pageSize);
    var hasMore = candidateRows.length > skipCount + pageSize;
    var pageRows = wmReadGroupedRowsFast_(sheet, pageRowNumbers, lastCol);
    var records = [];

    for (var i = 0; i < pageRows.length; i++) {
      var row = pageRows[i] || [];
      var actualStudentId = String(row[idxStudentId] || '').trim();
      if (wmStudentIdCompareKey_(actualStudentId) !== studentCompareKey) continue;

      var rawRecord = buildRecordObjectByHeaders_(headers, row);
      rawRecord.학생ID = actualStudentId;
      rawRecord.Set_ID = String(rawRecord.Set_ID || row[idxSetId] || '').trim();
      if (!rawRecord.Set_ID) continue;

      rawRecord.학생이름 = String(rawRecord.학생이름 || studentProfile.학생이름 || studentProfile.studentName || '').trim();
      rawRecord.학교 = String(rawRecord.학교 || studentProfile.학교 || studentProfile.school || '').trim();
      rawRecord.학년 = String(rawRecord.학년 || studentProfile.학년 || studentProfile.grade || '').trim();
      rawRecord.Class = String(rawRecord.Class || rawRecord['반'] || rawRecord['반명'] || studentProfile.Class || studentProfile.className || '').trim();
      rawRecord.교사명 = String(rawRecord.교사명 || studentProfile.교사명 || studentProfile.teacherName || '').trim();
      rawRecord.__displayRecordId = String(rawRecord['학습기록ID'] || '').trim();
      records.push(buildLearningMapDisplayRecord_(rawRecord));
    }

    records.sort(function(a, b) {
      return getLearningRecordSortTime_(b && b['학습날짜']) - getLearningRecordSortTime_(a && a['학습날짜']);
    });

    payload.records = records;
    payload.learningRecords = records;
    payload.recordsAll = records;
    payload.learningRecordDbDirect = true;
    payload.learningRecordDbPaged = true;
    payload.learningRecordDbPage = page;
    payload.learningRecordDbPageSize = pageSize;
    payload.learningRecordDbDirectCount = records.length;
    payload.learningRecordDbHasMore = hasMore;
    if (needTotalCount) payload.learningRecordDbSearchTotalCount = totalMatchedCount;
    payload.learningRecordDbFromDate = dateRange.fromDate;
    payload.learningRecordDbToDate = dateRange.toDate;
    payload.learningRecordDbCache = false;
  } catch (err) {
    payload.learningRecordDbDirect = false;
    payload.learningRecordDbDirectError = String(err && err.message ? err.message : err);
  }
  return payload;
}

/* WM_MAP_BOTTOM_RECORDS_DIRECT_API_20260707_V1
 * Map.html 하단 학습기록 전용 보강 API입니다.
 * 현재진행_DB 캐시/최근30건JSON 전달이 비어도 2.학습기록_DB를 학생ID 기준으로 직접 읽어 records를 반환합니다.
 */
function getLearningRecordsForMapData(studentId, page, pageSize, fromDate, toDate, searchFilterJson) {
  studentId = String(studentId || '').trim();
  /* WM_REAL_MAP_RECORD_PROFILE_DUPLICATE_READ_OFF_V1_20260812
   * 상세기록 호출에서 학생기본정보 시트를 다시 읽지 않습니다.
   * Map이 이미 가진 학생정보로 표시값을 보강하므로 학습기록_DB만 조회합니다. */
  var profile = {};
  var payload = {
    success: !!studentId,
    message: studentId ? '학습기록_DB 직접 조회 성공' : 'studentId가 없습니다.',
    studentId: studentId,
    records: [],
    learningRecords: []
  };
  if (!studentId) return payload;

  /* WM_MAP_RECORD_FULL_FETCH_BLOCK_V1_20260822
   * Map 학습기록 API는 전체조회 모드를 제공하지 않습니다.
   * pageSize가 없거나 0이어도 15개 페이지 조회로 강제합니다. */
  var safePageSize = Number(pageSize || 0) > 0 ? Number(pageSize) : 15;
  safePageSize = Math.max(1, Math.min(15, Math.floor(safePageSize)));

  return wmAttachLearningRecordDbPageForMap_(
    payload,
    studentId,
    profile,
    page,
    safePageSize,
    fromDate,
    toDate,
    searchFilterJson
  );
}

function getLearningMap(e) {
  var studentId = '';
  if (e && e.parameter && e.parameter.studentId) {
    studentId = String(e.parameter.studentId || '').trim();
  }

  return outputResult(e, buildLearningMapPayload_(studentId));
}

/* WM_GET_LEARNING_MAP_GOOGLE_SCRIPT_RUN_V1
 * Map.html이 HtmlService iframe 안에서 fetch CORS/redirect 문제로 Failed to fetch가 날 수 있으므로,
 * google.script.run으로 직접 호출 가능한 순수 JSON 객체 반환 함수를 제공합니다.
 */
function getLearningMapData(studentId) {
  return buildLearningMapPayload_(studentId);
}

/* WM_MAP_NEW_STUDENT_CURRENT_PROGRESS_INIT_20260820_V1
 * 신규학생은 첫 Map 진입 때만 1.학생관리_DB의 실제 학생정보/학습배정을 읽어
 * 8.현재진행_DB 학생 1행을 먼저 생성합니다.
 * 레벨/세트 기본값은 만들지 않으며 실제 학습배정 또는 현재세트가 확인될 때만 생성합니다. */
function wmEnsureInitialCurrentProgressForStudent_(studentId) {
  studentId = String(studentId || '').trim();
  if (!studentId) return {success:false, message:'studentId가 없습니다.'};

  var existing = wmGetCurrentProgressCacheForStudent_(studentId);
  if (existing) return {success:true, created:false, currentProgressCache:existing};

  var profile = getStudentBasicInfoForMap_(studentId) || {};
  var actualStudentId = String(profile.studentId || studentId).trim();
  var assignedLevel = extractLevelNumberForMap_(profile.learningAssign || '');
  var profileCurrentSet = normalizeCurrentSetForMap_(profile.currentSet || '', profile.learningAssign || '');
  var profileCurrentLevel = extractLevelNumberForMap_(profileCurrentSet || '');
  var initialLevel = profileCurrentLevel || assignedLevel;
  if (!actualStudentId || !profile.studentName || initialLevel < 3 || initialLevel > 13) {
    return {success:false, created:false, message:'학생관리_DB의 학생정보·현재세트·학습배정에서 실제 최초진행을 확인할 수 없습니다.'};
  }

  var initialPlainSet = profileCurrentSet || (initialLevel + '-1-1');
  var initialSetId = 'WM' + initialPlainSet;
  var sequence = getLevelSetSequence_(initialLevel) || [];
  if (!sequence.length || sequence.indexOf(initialPlainSet) < 0) {
    return {success:false, created:false, message:'학생관리_DB의 실제 현재세트/학습배정에 맞는 최초 세트를 확인할 수 없습니다.'};
  }

  var ss = getLmsSpreadsheet_();
  var sheet = ss && ss.getSheetByName('8.현재진행_DB');
  if (!sheet || sheet.getLastColumn() < 1) {
    return {success:false, created:false, message:'8.현재진행_DB를 찾을 수 없습니다.'};
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(h){
    return String(h || '').trim();
  });
  var idxStudent = headers.indexOf('학생ID');
  if (idxStudent < 0) return {success:false, created:false, message:'8.현재진행_DB 학생ID 컬럼이 없습니다.'};

  /* 학생관리_DB 조회 중 다른 저장경로가 현재진행 행을 만든 경우 중복생성을 피합니다. */
  var recheck = wmGetCurrentProgressCacheForStudent_(actualStudentId);
  if (recheck) return {success:true, created:false, currentProgressCache:recheck};

  var learningMode = profile.learningMode || buildDefaultLearningMode_();
  /* WM_CURRENT_PROGRESS_ORDER_TIME_MS_V1_20260821
   * 동시 완료 정렬 충돌을 최소화하기 위해 최종수정일을 밀리초까지 기록합니다.
   * 완전히 같은 시각이면 학생ID 오름차순을 2차 정렬키로 사용합니다. */
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  var initial = {
    학생ID: actualStudentId,
    학생이름: String(profile.studentName || '').trim(),
    학교: String(profile.school || '').trim(),
    학년: String(profile.grade || '').trim(),
    Class: String(profile.className || '').trim(),
    교사명: String(profile.teacherName || '').trim(),
    학습배정: String(profile.learningAssign || '').trim(),
    Set_ID: initialSetId,
    완료Step: '',
    현재Step: 'STEP1',
    기록Set_ID: initialSetId,
    최종수정일: nowText,
    학습기록ID: '',
    '1회차점수': '',
    '2회차점수': '',
    '3회차점수': '',
    최근점수: '',
    S4실루엣단계: String(learningMode.S4실루엣단계 || '').trim(),
    S6테스트모드: String(learningMode.S6테스트모드 || '').trim(),
    현재레벨: initialLevel + '레벨',
    레벨완료조건: String(learningMode.레벨완료조건 || '').trim(),
    레벨완료횟수: '0',
    순차완주세트: 0,
    순차완주회차: '0',
    완료세트수: 0,
    전체세트수: sequence.length,
    세트진행률: '0%',
    레벨진행률: '0%',
    진행률: '0%',
    레벨완료여부: 'N',
    현재세트: initialPlainSet,
    다음세트: '',
    최근30건JSON: '[]',
    캐시버전: 'v1',
    비고: ''
  };

  var row = headers.map(function(header){
    return initial[header] !== undefined ? initial[header] : '';
  });
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, headers.length).setValues([row]);

  var created = buildRecordObjectByHeaders_(headers, row);
  created.__rowNumber = 2;
  wmCachePutJson_(wmCacheKey_('CURRENT_PROGRESS_STUDENT', actualStudentId.toUpperCase()), created, 60);
  return {success:true, created:true, currentProgressCache:created};
}

/* WM_MAP_CURRENT_PROGRESS_CACHE_API_20260704_V1
 * Map.html이 직접 호출하는 현재진행_DB 전용 캐시 응답 함수입니다.
 * 성공/실패 모두 반드시 JSON 객체를 반환해서 학습맵 무한로딩을 끊습니다.
 */
function getLearningMapCacheData(studentId) {
  studentId = String(studentId || '').trim();

  if (!studentId) {
    return {
      success: false,
      message: 'studentId가 없습니다.',
      source: '8.현재진행_DB',
      studentId: ''
    };
  }

  try {
    var currentProgressCache = wmGetCurrentProgressCacheForStudent_(studentId);

    if (!currentProgressCache) {
      var initialProgressResult = wmEnsureInitialCurrentProgressForStudent_(studentId);
      if (initialProgressResult && initialProgressResult.success && initialProgressResult.currentProgressCache) {
        currentProgressCache = initialProgressResult.currentProgressCache;
      }
    }

    if (!currentProgressCache) {
      return {
        success:false,
        message:'현재진행_DB에서 학생 현재진행을 확인할 수 없습니다.',
        source:'8.현재진행_DB',
        studentId:studentId,
        currentProgressMissing:true,
        hasCurrentProgress:false,
        studentProfile:buildStudentProfileForLearningMap_(studentId),
        records:[],
        completedSetIds:[],
        levelHistoryLevels:[],
        setStatusMap:{}
      };
    }

    /* WM_MAP_CURRENT_PROGRESS_ONLY_FIRST_PROFILE_20260821_V1
     * 기존학생 첫 Map 화면은 8.현재진행_DB 학생 1행만 사용합니다.
     * 이름/학교/학년/Class/교사명/현재레벨/현재세트/학습조건은 현재진행_DB에서 바로 구성하고
     * 1.학생관리_DB를 다시 조회하지 않습니다. 신규학생의 최초 현재진행 행 생성 경로는 그대로 유지합니다. */
    var studentProfile = wmApplyCurrentProgressCacheToStudentProfile_({}, currentProgressCache);

    var currentProgressState = wmBuildCurrentProgressLearningStateForMap_(studentProfile, currentProgressCache);
    var payload = wmBuildFastLearningMapPayloadFromCurrentProgress_(studentId, studentProfile, currentProgressCache, currentProgressState);

    /* WM_SPEED_MAP_FAST_RESPONSE_V1_20260710
     * 정상 현재진행_DB 경로에서는 2.학습기록_DB 전체 직접 조회를 기다리지 않습니다.
     * Map.html이 getLearningRecordsForMapData를 병렬 호출하므로 맵 응답을 먼저 반환합니다.
     * 현재진행_DB가 없거나 오류인 fallback 경로의 기존 학습기록 조회는 그대로 유지합니다.
     */
    if (payload) {
      payload.recordsDeferred = true;
      payload.records = Array.isArray(payload.records) ? payload.records : [];
      payload.learningRecords = Array.isArray(payload.learningRecords) ? payload.learningRecords : payload.records;
    }

    if (!payload) {
      return {
        success: false,
        message: '현재진행_DB 응답 payload 생성 실패',
        source: '8.현재진행_DB',
        studentId: studentId,
        currentProgressCache: currentProgressCache,
        records: [],
        completedSetIds: [],
        levelHistoryLevels: [],
        setStatusMap: {}
      };
    }

    payload.currentProgressCache = currentProgressCache;
    payload.currentProgress = currentProgressCache;
    payload.currentProgressRow = currentProgressCache;
    payload.progressCache = currentProgressCache;
    payload.mapCache = currentProgressCache;
    payload.hasCurrentProgress = true;
    payload.source = '8.현재진행_DB';
    payload.cacheOnly = true;
    payload.message = payload.message || '현재진행_DB 캐시 조회 성공';
    return payload;
  } catch (err) {
    return {
      success:false,
      message:'getLearningMapCacheData 오류',
      source:'8.현재진행_DB',
      studentId:studentId,
      error:String(err && err.message ? err.message : err),
      currentProgressError:String(err && err.message ? err.message : err),
      hasCurrentProgress:false,
      records:[],
      completedSetIds:[],
      levelHistoryLevels:[],
      setStatusMap:{}
    };
  }
}

function buildLearningMapPayload_(studentId) {
  /* WM_MAP_CURRENT_PROGRESS_FIRST_PAYLOAD_20260703_V1
   * 학습맵은 8.현재진행_DB 학생 1행으로 먼저 완성합니다.
   * 2.학습기록_DB 전체조회는 현재진행_DB가 없을 때만 fallback으로 실행합니다.
   */
  studentId = String(studentId || '').trim();

  if (!studentId) {
    return {
      success: false,
      message: 'studentId가 없습니다.'
    };
  }

  var currentProgressCache = wmGetCurrentProgressCacheForStudent_(studentId);

  if (currentProgressCache) {
    var fastStudentProfile = buildStudentProfileForLearningMap_(studentId);
    fastStudentProfile = wmApplyCurrentProgressCacheToStudentProfile_(fastStudentProfile, currentProgressCache);

    var fastCurrentProgressState = wmBuildCurrentProgressLearningStateForMap_(fastStudentProfile, currentProgressCache);
    var fastPayload = wmBuildFastLearningMapPayloadFromCurrentProgress_(studentId, fastStudentProfile, currentProgressCache, fastCurrentProgressState);
    fastPayload = wmAttachLearningRecordDbRecordsForMap_(fastPayload, studentId, fastStudentProfile);
    if (fastPayload) return fastPayload;
  }

  var ss = getLmsSpreadsheet_();

  if (!ss) {
    return {
      success: false,
      message: '스프레드시트 연결 실패'
    };
  }

  var studentProfile = buildStudentProfileForLearningMap_(studentId);
  currentProgressCache = currentProgressCache || wmGetCurrentProgressCacheForStudent_(studentId);
  studentProfile = wmApplyCurrentProgressCacheToStudentProfile_(studentProfile, currentProgressCache);
  var currentProgressState = wmBuildCurrentProgressLearningStateForMap_(studentProfile, currentProgressCache);
  var progressCurrentSet = currentProgressState.currentSet || wmGetCurrentProgressSetForStudent_(studentId);

  if (progressCurrentSet) {
    studentProfile.현재세트 = progressCurrentSet;
    studentProfile.currentSet = progressCurrentSet;
  }

  var fallbackFastPayload = wmBuildFastLearningMapPayloadFromCurrentProgress_(studentId, studentProfile, currentProgressCache, currentProgressState);
  fallbackFastPayload = wmAttachLearningRecordDbRecordsForMap_(fallbackFastPayload, studentId, studentProfile);
  if (fallbackFastPayload) return fallbackFastPayload;

  var sheet = ss.getSheetByName('2.학습기록_DB');

  if (!sheet) {
    return {
      success: false,
      message: '2.학습기록_DB 시트를 찾을 수 없습니다.'
    };
  }

  var values = sheet.getDataRange().getDisplayValues();

  if (!values || values.length < 2) {
    return {
      success: true,
      message: '학습기록이 없습니다.',
      studentId: studentId,
      studentProfile: studentProfile,
      currentProgressCache: currentProgressCache,
      records: [],
      completedSetIds: [],
      officialLevelProgress: currentProgressState.officialLevelProgress,
      currentLearningProgress: currentProgressState.currentLearningProgress,
      levelHistoryLevels: [],
      setStatusMap: currentProgressState.setStatusMap
    };
  }

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || '').trim());
  }

  var idxRecordId = headers.indexOf('학습기록ID');
  var idxDate = headers.indexOf('학습날짜');
  var idxStudentId = headers.indexOf('학생ID');
  var idxStudentName = headers.indexOf('학생이름');
  var idxSetId = headers.indexOf('Set_ID');
  var idxScore = headers.indexOf('점수');
  var idxTotalTime = headers.indexOf('총소요시간');
  var idxTestTime = headers.indexOf('Test_총시간');
  var idxStatus = headers.indexOf('완료상태');

  if (idxStudentId === -1) {
    return {
      success: false,
      message: '2.학습기록_DB에서 학생ID 컬럼을 찾을 수 없습니다.'
    };
  }

  if (idxSetId === -1) {
    return {
      success: false,
      message: '2.학습기록_DB에서 Set_ID 컬럼을 찾을 수 없습니다.'
    };
  }

  if (idxStatus === -1) {
    return {
      success: false,
      message: '2.학습기록_DB에서 완료상태 컬럼을 찾을 수 없습니다.'
    };
  }

  var records = [];
  var setStatusMap = {};
  var completedSetIds = [];
  var completedLearningLogs = [];
  var recordIdDisplayState = buildLearningMapRecordIdDisplayState_(headers, values);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowStudentId = String(row[idxStudentId] || '').trim();

    if (wmStudentIdCompareKey_(rowStudentId) !== wmStudentIdCompareKey_(studentId)) {
      continue;
    }

    var setId = String(row[idxSetId] || '').trim();
    var completeStatus = String(row[idxStatus] || '').trim();

    if (!setId) {
      continue;
    }

    var rawRecord = buildRecordObjectByHeaders_(headers, row);
    rawRecord.학생ID = rowStudentId;
    rawRecord.Set_ID = setId;
    rawRecord.완료상태 = completeStatus;
    rawRecord.완료 = String(rawRecord.완료 || '').trim();

    /* WM_LEARNING_MAP_PROFILE_FALLBACK_V2
     * 기존 학습기록 행에 학교/학년/Class/교사명이 비어 있어도
     * 학습맵 표시용 records에는 1.학생관리_DB 기준 studentProfile 값으로 보강합니다.
     * DB 원본 행을 임의 수정하지 않고, 화면/API 응답값만 보강합니다.
     */
    rawRecord.학교 = String(rawRecord.학교 || studentProfile.학교 || '').trim();
    rawRecord.학년 = String(rawRecord.학년 || studentProfile.학년 || '').trim();
    rawRecord.Class = String(rawRecord.Class || rawRecord['반'] || rawRecord['반명'] || studentProfile.Class || '').trim();
    rawRecord.교사명 = String(rawRecord.교사명 || studentProfile.교사명 || '').trim();
    rawRecord.__displayRecordId = normalizeLearningMapRecordIdForDisplay_(rawRecord['학습기록ID'], rawRecord['학습날짜'], recordIdDisplayState);

    /* WM_LEARNING_MAP_RECORD_PAYLOAD_FILTER_V3
     * 학습맵 화면에는 운영에 필요한 학습 진행/결과 값만 내려줍니다.
     * 학습기록ID와 기기정보는 화면 확인에 필요하므로 내려줍니다.
     * 기기정보는 PC/Mobile/Tap으로 정리하고, 완료상태는 최종 완료일 때만 '완료'로 내려줍니다.
     */
    var record = buildLearningMapDisplayRecord_(rawRecord);
    records.push(record);

    if (!setStatusMap[setId]) {
      setStatusMap[setId] = {
        setId: setId,
        repeatCount: 0,
        officialRound: 0,
        latestStatus: '',
        latestScore: '',
        latestDate: '',
        currentProgressStep: 'STEP1',
        lastCompletedStep: '',
        completedStepCount: 0,
        progressPercent: 0,
        progressRank: 0
      };
    }

    var rowProgress = deriveProgressFromLearningRecord_(rawRecord);
    var rowProgressRank = getLearningProgressRankForMap_(rowProgress.currentProgressStep);
    var rowProgressPercent = getLearningMapProgressPercent_(rowProgress);

    if (isCompleteStatus(completeStatus)) {
      setStatusMap[setId].repeatCount += 1;
      setStatusMap[setId].officialRound = Math.max(Number(setStatusMap[setId].officialRound || 0), wmGetOfficialRoundFromRow_(headers, row, completeStatus));
      setStatusMap[setId].latestStatus = completeStatus;
      setStatusMap[setId].latestScore = record.한영주관식 || record.점수;
      setStatusMap[setId].latestDate = record.학습날짜;
      setStatusMap[setId].currentProgressStep = 'COMPLETE';
      setStatusMap[setId].lastCompletedStep = 'TEST';
      setStatusMap[setId].completedStepCount = 6;
      setStatusMap[setId].progressPercent = 100;
      setStatusMap[setId].progressRank = 6;

      if (completedSetIds.indexOf(setId) === -1) {
        completedSetIds.push(setId);
      }

      completedLearningLogs.push({
        setId: setId,
        learningDate: record.학습날짜 || rawRecord['학습날짜'] || '',
        sortTime: getLearningRecordSortTime_(record.학습날짜 || rawRecord['학습날짜'] || ''),
        rowIndex: i
      });
    } else if (rowProgressPercent > 0 && rowProgressRank >= setStatusMap[setId].progressRank) {
      /* WM_MAP_IN_PROGRESS_PERCENT_V1
       * 한 Step이라도 완료된 미완료 세트는 학습중 %로 표시합니다.
       * 세트진행률 공식: 완료 Step 수 / 6단계 * 100, 소수점 버림.
       */
      setStatusMap[setId].latestStatus = '학습중';
      setStatusMap[setId].latestDate = record.학습날짜;
      setStatusMap[setId].currentProgressStep = rowProgress.currentProgressStep;
      setStatusMap[setId].lastCompletedStep = rowProgress.lastCompletedStep;
      setStatusMap[setId].completedStepCount = rowProgress.completedSteps.length;
      setStatusMap[setId].progressPercent = rowProgressPercent;
      setStatusMap[setId].progressRank = rowProgressRank;
    }
  }

  records.sort(function(a, b) {
    return getLearningRecordSortTime_(b && b['학습날짜']) - getLearningRecordSortTime_(a && a['학습날짜']);
  });

  var officialLevelProgress = buildOfficialLevelProgressForLearningMap_(studentProfile, completedLearningLogs, setStatusMap);
  officialLevelProgress = wmApplyCurrentProgressCacheToOfficialLevelProgress_(officialLevelProgress, currentProgressCache);

  filterLockedSetProgressForLearningMap_(setStatusMap, completedSetIds, studentProfile, officialLevelProgress);

  var currentLearningProgress = buildCurrentLearningProgressForMap_(setStatusMap, completedSetIds, studentProfile, officialLevelProgress);

  applyOfficialLevelProgressToSetStatusMap_(setStatusMap, officialLevelProgress);

  addLearningMapSetStatusAliases_(setStatusMap, completedSetIds);

  Object.keys(currentProgressState.setStatusMap || {}).forEach(function(key) {
    if (!setStatusMap[key]) {
      setStatusMap[key] = currentProgressState.setStatusMap[key];
    }
  });

  if (currentProgressState.officialLevelProgress) {
    officialLevelProgress = wmApplyCurrentProgressCacheToOfficialLevelProgress_(officialLevelProgress || currentProgressState.officialLevelProgress, currentProgressCache);
  }

  return {
    success: true,
    message: '학습기록 조회 성공',
    studentId: studentId,
    studentProfile: studentProfile,
    currentProgressCache: currentProgressCache,
    records: records,
    completedSetIds: completedSetIds,
    officialLevelProgress: officialLevelProgress,
    currentLearningProgress: currentLearningProgress,
    levelHistoryLevels: buildLevelHistoryLevelsForLearningMap_(records, completedSetIds, setStatusMap, studentProfile, officialLevelProgress),
    setStatusMap: setStatusMap
  };
}


function buildOfficialLevelProgressForLearningMap_(studentProfile, completedLearningLogs, setStatusMap) {
  /* WM_OFFICIAL_LEVEL_PROGRESS_INTEGRATION_20260616_V1
   * 실제 학습맵에 연결되는 공식 순차완주 상태입니다.
   * 완료 기록을 날짜순으로 읽어 현재 라운드의 기대 세트와 일치할 때만 공식 회차로 인정합니다.
   * 즉, 레벨 전체 1회 완주 전의 2회차 재학습은 공식 2회차로 이월 인정하지 않습니다.
   */
  var learningAssign = (studentProfile && studentProfile['학습배정']) || '4레벨';
  var level = extractLevelNumberForMap_(learningAssign);
  var learningMode = (studentProfile && studentProfile['학습모드']) || buildDefaultLearningMode_();
  var goal = buildLevelTargetRoundGoalFromLearningMode_(learningMode);
  var targetRounds = Number(goal.목표회차 || 1);
  var sequence = getLevelSetSequence_(level);
  var officialCounts = {};
  var officialEvents = [];
  var currentRound = 1;
  var currentIndex = 0;

  if (!sequence.length) {
    return {
      level: level,
      목표회차: targetRounds,
      공식완료회차: 0,
      현재공식회차: 1,
      다음공식세트: buildInitialSetIdFromLearningAssign(learningAssign),
      레벨완료: false,
      다음레벨자동열림: false,
      다음레벨첫세트: '',
      완료세트수: 0,
      전체세트수: 0,
      세트진행률: 0,
      레벨진행률: 0,
      진행률: 0,
      progressPercent: 0,
      levelProgressPercent: 0,
      stepProgressPercent: 0,
      officialCounts: officialCounts,
      officialEvents: officialEvents,
      goal: goal
    };
  }

  var logs = Array.isArray(completedLearningLogs) ? completedLearningLogs.slice() : [];
  logs.sort(function(a, b) {
    var at = Number(a && a.sortTime || 0);
    var bt = Number(b && b.sortTime || 0);
    if (at !== bt) return at - bt;
    return Number(a && a.rowIndex || 0) - Number(b && b.rowIndex || 0);
  });

  for (var i = 0; i < logs.length; i++) {
    if (currentRound > targetRounds) break;

    var plainSetId = normalizeLevelPlainSetId_(logs[i] && logs[i].setId);
    if (!plainSetId) continue;
    if (plainSetId.indexOf(String(level) + '-') !== 0) continue;

    var expectedSetId = sequence[currentIndex];
    if (plainSetId !== expectedSetId) {
      continue;
    }

    officialCounts[plainSetId] = Number(officialCounts[plainSetId] || 0) + 1;
    officialEvents.push({
      setId: plainSetId,
      officialRound: currentRound,
      setIndex: currentIndex + 1,
      learningDate: String((logs[i] && logs[i].learningDate) || '')
    });

    currentIndex += 1;
    if (currentIndex >= sequence.length) {
      currentRound += 1;
      currentIndex = 0;
    }
  }

  var completedRounds = Math.max(0, currentRound - 1);
  if (completedRounds > targetRounds) completedRounds = targetRounds;

  var isComplete = completedRounds >= targetRounds;
  var nextOfficialSet = isComplete ? '' : sequence[currentIndex];
  var nextLevelFirstSet = isComplete ? getNextLevelFirstSetId_(level) : '';

  var officialCompletedSetCount = 0;
  for (var oc = 0; oc < sequence.length; oc++) {
    if (Number(officialCounts[sequence[oc]] || 0) > 0) officialCompletedSetCount += 1;
  }
  var officialProgressPercent = sequence.length > 0
    ? Math.round((officialCompletedSetCount / sequence.length) * 100)
    : 0;

  return {
    level: level,
    목표회차: targetRounds,
    공식완료회차: completedRounds,
    현재공식회차: isComplete ? targetRounds : currentRound,
    완료세트수: officialCompletedSetCount,
    전체세트수: sequence.length,
    세트진행률: 0,
    레벨진행률: officialProgressPercent,
    진행률: officialProgressPercent,
    progressPercent: officialProgressPercent,
    levelProgressPercent: officialProgressPercent,
    stepProgressPercent: 0,
    다음공식세트: nextOfficialSet,
    현재회차세트순번: isComplete ? sequence.length : currentIndex + 1,
    레벨완료: isComplete,
    다음레벨자동열림: !!nextLevelFirstSet,
    다음레벨: nextLevelFirstSet ? level + 1 : '',
    다음레벨첫세트: nextLevelFirstSet,
    officialCounts: officialCounts,
    officialEvents: officialEvents,
    goal: goal
  };
}

function buildSetSequentialDisplayState_(officialCount, targetRounds, isLevelComplete, hasInProgress) {
  /* WM_SET_SEQUENTIAL_COLOR_RULE_20260616_V1
   * 세트순차완주 표시 공식입니다.
   * - 실제 복습 횟수가 아니라 공식 순차완주 회차만 화면 색상/문구에 반영합니다.
   * - 레벨완료조건이 1회이면 2회/3회 복습 기록이 있어도 파란색(1회 완료)까지만 표시합니다.
   * - 레벨완료조건이 2회이면 3회 복습 기록이 있어도 주황색(2회 완료)까지만 표시합니다.
   * - 레벨완료조건이 3회이면 3회 완료를 보라색으로 표시합니다.
   * - 레벨완료 상태는 별도 초록색 상태로 보강하되, 블록 3줄 문구는 n회 완료 형식을 유지합니다.
   */
  var target = Number(targetRounds || 1);
  if (!isFinite(target) || target < 1) target = 1;
  if (target > 3) target = 3;

  var official = Number(officialCount || 0);
  if (!isFinite(official) || official < 0) official = 0;

  var displayRound = Math.min(Math.floor(official), target);
  var state = {
    displayRound: displayRound,
    displayText: '',
    colorState: '',
    colorKey: '',
    isLearning: false,
    isComplete: !!isLevelComplete
  };

  if (displayRound >= 1) {
    state.displayText = displayRound + '회 완료';
    state.colorState = isLevelComplete ? 'LEVEL_COMPLETE' : ('ROUND_' + displayRound);
    state.colorKey = isLevelComplete ? 'green' : (displayRound === 1 ? 'blue' : (displayRound === 2 ? 'orange' : 'purple'));
    return state;
  }

  if (hasInProgress) {
    state.displayText = '학습중';
    state.colorState = 'IN_PROGRESS';
    state.colorKey = 'yellow';
    state.isLearning = true;
    return state;
  }

  state.displayText = '';
  state.colorState = 'LOCKED_OR_NOT_STARTED';
  state.colorKey = 'locked';
  return state;
}

function applyOfficialLevelProgressToSetStatusMap_(setStatusMap, officialLevelProgress) {
  /* WM_OFFICIAL_STATUS_PAYLOAD_20260616_V2
   * Map.html이 활용할 수 있도록 각 세트 상태에 공식회차/공식다음세트/세트순차완주 색상 기준을 보강합니다.
   */
  if (!setStatusMap || typeof setStatusMap !== 'object' || !officialLevelProgress) return;

  var nextOfficial = normalizeStudySetIdForCompare_(officialLevelProgress.다음공식세트 || officialLevelProgress.다음레벨첫세트 || '');
  var officialCounts = officialLevelProgress.officialCounts || {};
  var targetRounds = Number(officialLevelProgress.목표회차 || 1);
  var currentLevel = Number(officialLevelProgress.level || 0);

  Object.keys(setStatusMap).forEach(function(key) {
    var status = setStatusMap[key] || {};
    var setId = normalizeLevelPlainSetId_(status.setId || key);
    if (!setId) return;

    var setLevel = Number(setId.split('-')[0] || 0);
    var officialCount = Number(officialCounts[setId] || 0);
    if (!isFinite(officialCount) || officialCount < 0) officialCount = 0;

    var hasInProgress =
      Number(status.progressPercent || 0) > 0 ||
      Number(status.completedStepCount || 0) > 0 ||
      String(status.latestStatus || '').trim() === '학습중';

    var displayState = buildSetSequentialDisplayState_(
      officialCount,
      targetRounds,
      !!officialLevelProgress.레벨완료 && setLevel === currentLevel,
      hasInProgress
    );

    status.rawRepeatCount = Number(status.repeatCount || 0);
    status.repeatCountOfficial = officialCount;
    status.공식완료회차 = officialCount;
    status.공식표시회차 = displayState.displayRound;
    status.공식표시문구 = displayState.displayText;
    status.세트순차완주상태 = displayState.colorState;
    status.세트색상기준 = displayState.colorKey;

    if (setLevel === currentLevel) {
      /* Map.html 기존 색상 로직이 repeatCount를 보더라도 공식 표시회차 기준으로만 색이 변하도록 고정합니다. */
      status.repeatCount = displayState.displayRound;
      if (displayState.displayRound > 0) {
        status.latestStatus = '완료';
      } else if (displayState.isLearning) {
        status.latestStatus = '학습중';
      } else {
        status.latestStatus = '';
      }
    }

    status.목표회차 = Number(officialLevelProgress.목표회차 || 0);
    status.현재공식회차 = Number(officialLevelProgress.현재공식회차 || 1);
    status.공식완료세트수 = Number(officialLevelProgress.완료세트수 || 0);
    status.공식전체세트수 = Number(officialLevelProgress.전체세트수 || 0);
    status.세트진행률 = Number(officialLevelProgress.세트진행률 || status.progressPercent || 0);
    status.레벨진행률 = Number(officialLevelProgress.레벨진행률 || officialLevelProgress.progressPercent || officialLevelProgress.진행률 || 0);
    status.공식진행률 = String(status.레벨진행률 || 0) + '%';
    status.officialProgressPercent = Number(status.레벨진행률 || 0);
    status.isOfficialNextSet = normalizeStudySetIdForCompare_(setId) === nextOfficial;
    status.레벨완료 = !!officialLevelProgress.레벨완료 && setLevel === currentLevel;
    status.다음레벨자동열림 = !!officialLevelProgress.다음레벨자동열림;
  });
}

function filterLockedSetProgressForLearningMap_(setStatusMap, completedSetIds, studentProfile, officialLevelProgress) {
  /* WM_DB_PROGRESS_PRESERVE_FOR_MAP_20260616_V1
   * 학습기록_DB에 Step 진행 기록이 있으면 상단 학습맵 블럭 %에 반드시 반영합니다.
   * 기존 잠금 필터가 현재 공식 세트가 아니라고 판단해도,
   * DB에 progressPercent/completedStepCount/latestStatus=학습중 값이 있으면 0%로 지우지 않습니다.
   */
  if (!setStatusMap || typeof setStatusMap !== 'object') {
    return;
  }

  var unlockedSetId = getCurrentUnlockedSetIdForMap_(completedSetIds, studentProfile, setStatusMap, officialLevelProgress);

  Object.keys(setStatusMap).forEach(function(key) {
    var status = setStatusMap[key] || {};
    var setId = String(status.setId || key || '').trim();

    if (isCompleteStatus(status.latestStatus)) {
      return;
    }

    var hasDbProgress =
      Number(status.progressPercent || 0) > 0 ||
      Number(status.completedStepCount || 0) > 0 ||
      String(status.latestStatus || '').trim() === '학습중';

    if (hasDbProgress) {
      if (!status.latestStatus) status.latestStatus = '학습중';
      return;
    }

    if (normalizeStudySetIdForCompare_(setId) === normalizeStudySetIdForCompare_(unlockedSetId)) {
      return;
    }

    status.latestStatus = '';
    status.latestScore = '';
    status.latestDate = '';
    status.currentProgressStep = 'STEP1';
    status.lastCompletedStep = '';
    status.completedStepCount = 0;
    status.progressPercent = 0;
    status.progressRank = 0;
  });
}

function findLatestInProgressSetStatusForMap_(setStatusMap) {
  if (!setStatusMap || typeof setStatusMap !== 'object') return null;

  var best = null;
  Object.keys(setStatusMap).forEach(function(key) {
    var status = setStatusMap[key] || {};
    if (isCompleteStatus(status.latestStatus)) return;

    var percent = Number(status.progressPercent || 0);
    var stepCount = Number(status.completedStepCount || 0);
    if (percent <= 0 && stepCount <= 0) return;

    var sortTime = getLearningRecordSortTime_(status.latestDate || '');
    if (!best || sortTime >= best.sortTime) {
      best = {
        setId: String(status.setId || key || '').trim(),
        status: status,
        sortTime: sortTime
      };
    }
  });

  return best;
}


function buildCurrentLearningProgressForMap_(setStatusMap, completedSetIds, studentProfile, officialLevelProgress) {
  /* WM_CURRENT_BLOCK_PROGRESS_PAYLOAD_V2
   * 현재 열림 세트는 단순 1회 완료 기준이 아니라 레벨완료조건/추가회차/순차완주 공식 엔진 기준으로 계산합니다.
   */
  var currentSetId = getCurrentUnlockedSetIdForMap_(completedSetIds, studentProfile, setStatusMap, officialLevelProgress);
  var status = findSetStatusBySetId_(setStatusMap, currentSetId) || {};

  /* WM_CURRENT_PROGRESS_DB_FALLBACK_20260616_V1
   * 현재 열림 세트 계산값이 0%여도 DB에 학습중 기록이 있는 세트가 있으면 그 값을 우선 표시합니다.
   */
  if (Number(status.progressPercent || 0) <= 0) {
    var active = findLatestInProgressSetStatusForMap_(setStatusMap);
    if (active && Number(active.status.progressPercent || 0) > 0) {
      currentSetId = active.setId;
      status = active.status;
    }
  }

  var official = officialLevelProgress || buildOfficialLevelProgressForLearningMap_(studentProfile, [], setStatusMap);

  var officialProgressPercent = Number(official.레벨진행률 || official.progressPercent || official.진행률 || 0);
  var stepProgressPercent = Number(official.세트진행률 || status.progressPercent || 0);

  return {
    setId: currentSetId,
    progressPercent: officialProgressPercent,
    세트진행률: stepProgressPercent,
    레벨진행률: officialProgressPercent,
    stepProgressPercent: stepProgressPercent,
    levelProgressPercent: officialProgressPercent,
    currentProgressStep: String(status.currentProgressStep || 'STEP1'),
    lastCompletedStep: String(status.lastCompletedStep || ''),
    completedStepCount: Number(status.completedStepCount || 0),
    목표회차: Number(official.목표회차 || 0),
    공식완료회차: Number(official.공식완료회차 || 0),
    현재공식회차: Number(official.현재공식회차 || 1),
    다음공식세트: String(official.다음공식세트 || currentSetId || ''),
    레벨완료: !!official.레벨완료,
    다음레벨자동열림: !!official.다음레벨자동열림,
    다음레벨첫세트: String(official.다음레벨첫세트 || '')
  };
}


function findSetStatusBySetId_(setStatusMap, setId) {
  if (!setStatusMap || typeof setStatusMap !== 'object') {
    return null;
  }

  var wanted = normalizeStudySetIdForCompare_(setId);
  var keys = Object.keys(setStatusMap);
  for (var i = 0; i < keys.length; i++) {
    var item = setStatusMap[keys[i]] || {};
    var itemId = String(item.setId || keys[i] || '').trim();
    if (normalizeStudySetIdForCompare_(itemId) === wanted) {
      return item;
    }
  }

  return null;
}

function getCurrentUnlockedSetIdForMap_(completedSetIds, studentProfile, setStatusMap, officialLevelProgress) {
  /* WM_CURRENT_UNLOCKED_SET_SEQUENTIAL_V1
   * 레벨완료조건(1/2/3/추가) + 순차완주 공식 기준으로 현재 열릴 세트를 반환합니다.
   * 레벨이 목표회차까지 완료되면 다음 레벨 첫 세트를 반환합니다.
   */
  var official = officialLevelProgress || buildOfficialLevelProgressForLearningMap_(studentProfile, [], setStatusMap);

  if (official && official.다음레벨자동열림 && official.다음레벨첫세트) {
    return String(official.다음레벨첫세트 || '');
  }

  if (official && official.다음공식세트) {
    return String(official.다음공식세트 || '');
  }

  var startSetId = buildInitialSetIdFromLearningAssign((studentProfile && studentProfile['학습배정']) || '4레벨');
  var current = parseLearningMapPlainSetId_(startSetId);

  if (!current) {
    return startSetId;
  }

  var completedMap = {};
  if (Array.isArray(completedSetIds)) {
    completedSetIds.forEach(function(id) {
      completedMap[normalizeStudySetIdForCompare_(id)] = true;
    });
  }

  var guard = 0;
  while (guard < 500 && completedMap[formatLearningMapPlainSetId_(current)]) {
    current = getNextLearningMapSetIdParts_(current);
    guard += 1;
  }

  return formatLearningMapPlainSetId_(current);
}



function wmSyncStudentCurrentSetAfterOfficialCompletion_(studentId, completeStatus, prefetchedNextSetId) {
  /* WM_STUDENT_CURRENT_SET_OFFICIAL_SYNC_20260616_V1
   * 학습 완료 저장 후 1.학생관리_DB의 현재세트를 공식 순차완주 엔진 기준으로 동기화합니다.
   * 기준: 학생관리_DB 학습배정 / 레벨완료조건 / 레벨회차추가.
   * 목표회차 달성 전: 다음공식세트
   * 목표회차 달성 후: 다음레벨첫세트
   * 미완료 저장/Step 저장은 현재세트를 변경하지 않습니다.
   */
  studentId = String(studentId || '').trim().toUpperCase();
  if (!studentId || !isCompleteStatus(completeStatus)) {
    return { success: false, skipped: true, reason: 'not_completed' };
  }

  var ss = getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  if (!studentSheet) {
    return { success: false, message: '1.학생관리_DB 시트를 찾을 수 없습니다.' };
  }

  var lastCol = studentSheet.getLastColumn();
  if (lastCol < 1 || studentSheet.getLastRow() < 2) {
    return { success: false, message: '학생관리_DB 데이터가 없습니다.' };
  }

  var headers = studentSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  var idxStudentId = headers.indexOf('학생ID');
  var idxCurrentSet = headers.indexOf('현재세트');

  if (idxStudentId < 0 || idxCurrentSet < 0) {
    return { success: false, message: '학생ID/현재세트 컬럼을 찾을 수 없습니다.' };
  }

  var studentRows = wmGetStudentRowNumbersFast_(studentSheet, headers, studentId);
  var targetRow = studentRows.length ? studentRows[0] : -1;

  if (targetRow < 2) {
    return { success: false, message: '학생ID를 찾을 수 없습니다: ' + studentId };
  }

  var nextSetId = String(prefetchedNextSetId || '').trim();
  var official = {};
  var progress = {};

  /* 최종 저장에서 이미 계산한 다음세트가 있으면 학습맵 전체 payload를 다시 만들지 않습니다. */
  if (!nextSetId) {
    var payload = buildLearningMapPayload_(studentId);
    if (!payload || !payload.success) {
      return { success: false, message: '공식 진행 상태 계산 실패', payloadMessage: payload && payload.message };
    }
    official = payload.officialLevelProgress || {};
    progress = payload.currentLearningProgress || {};
  }

  /* WM_NEXT_LEVEL_UNLOCK_DB_CONDITION_GUARD_20260616_V1
   * 다음 레벨 첫 세트는 학생관리_DB의 레벨완료조건을 기준으로
   * 세트순차완주 목표회차를 실제 달성한 경우에만 현재세트로 반영합니다.
   * 마지막 세트를 1회 완료했더라도 목표회차 미달이면 다음레벨첫세트는 사용하지 않고,
   * 같은 레벨의 다음공식세트로 되돌립니다.
   */
  if (!nextSetId && official && official.레벨완료 === true && official.다음레벨자동열림 === true && official.다음레벨첫세트) {
    nextSetId = String(official.다음레벨첫세트 || '').trim();
  } else if (!nextSetId && official && official.다음공식세트) {
    nextSetId = String(official.다음공식세트 || '').trim();
  } else if (!nextSetId) {
    nextSetId = String(progress.setId || '').trim();
  }

  if (!nextSetId) {
    return { success: true, skipped: true, reason: 'no_next_set', officialLevelProgress: official };
  }

  studentSheet.getRange(targetRow, idxCurrentSet + 1).setValue(nextSetId);

  return {
    success: true,
    studentId: studentId,
    currentSet: nextSetId,
    목표회차: Number(official.목표회차 || 0),
    공식완료회차: Number(official.공식완료회차 || 0),
    레벨완료: !!official.레벨완료,
    다음레벨자동열림: !!official.다음레벨자동열림
  };
}

function parseLearningMapPlainSetId_(setId) {
  var text = normalizeStudySetIdForCompare_(setId);
  var match = text.match(/^(\d+)-(\d+)-(\d+)$/);
  if (!match) {
    return null;
  }
  return {
    level: Number(match[1]),
    part: Number(match[2]),
    set: Number(match[3])
  };
}

function formatLearningMapPlainSetId_(parts) {
  if (!parts) {
    return '';
  }
  return String(parts.level) + '-' + String(parts.part) + '-' + String(parts.set);
}

function getNextLearningMapSetIdParts_(parts) {
  var level = Number(parts && parts.level || 0);
  var part = Number(parts && parts.part || 1);
  var set = Number(parts && parts.set || 1) + 1;

  if (!level) {
    return { level: 3, part: 1, set: 1 };
  }

  var partCounts = getLearningMapLevelPartSetCounts_(level);
  if (!partCounts.length) {
    return { level: level, part: part, set: set };
  }

  var maxSet = Number(partCounts[part - 1] || 0);

  if (!maxSet || set > maxSet) {
    set = 1;
    part += 1;
  }

  if (part > partCounts.length) {
    level += 1;
    part = 1;
    set = 1;

    if (level > 13) {
      level = 13;
      part = partCounts.length;
      set = Number(partCounts[partCounts.length - 1] || 1);
    }
  }

  return { level: level, part: part, set: set };
}

function getLearningMapMaxPartByLevel_(level) {
  return getLearningMapLevelPartSetCounts_(level).length;
}
function getLearningMapLevelPartSetCounts_(level) {
  var n = Number(level || 0);

  if (n === 3) return [10, 10, 2];
  if (n === 4) return [10, 10, 2];
  if (n === 5) return [10, 10, 3];
  if (n === 6) return [10, 10, 3];
  if (n >= 7 && n <= 9) return [10, 10, 10, 10];
  if (n >= 10 && n <= 13) return [10, 10, 10, 10, 10];

  return [];
}
function getLearningRecordSortTime_(value) {
  var text = String(value || '').trim();
  if (!text) {
    return 0;
  }

  var normalized = text.replace(/\./g, '-').replace(/\//g, '-');
  var date = new Date(normalized);
  if (!isNaN(date.getTime())) {
    return date.getTime();
  }

  var match = text.match(/(\d{4})[-.\/년\s]+(\d{1,2})[-.\/월\s]+(\d{1,2})(?:[^0-9]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    ).getTime();
  }

  return 0;
}

function getLearningMapProgressPercent_(progress) {
  /* WM_PROGRESS_PERCENT_6_STEP_V1
   * 학습맵 세트진행률 공식: 완료 Step 수 / 6단계 * 100, 소수점 버림.
   */
  var completedSteps = progress && Array.isArray(progress.completedSteps) ? progress.completedSteps : [];
  var count = completedSteps.length;
  if (progress && String(progress.currentProgressStep || '') === 'COMPLETE') {
    count = 6;
  }
  return Math.floor((Math.min(Math.max(count, 0), 6) / 6) * 100);
}

function getLearningProgressRankForMap_(step) {
  var text = String(step || '').trim().toUpperCase();
  if (text === 'COMPLETE') return 6;
  if (text === 'TEST' || text === 'STEP6') return 5;
  if (text === 'STEP5') return 4;
  if (text === 'BEFORE_TEST') return 4;
  if (text === 'STEP4') return 3;
  if (text === 'STEP3') return 2;
  if (text === 'STEP2') return 1;
  return 0;
}

function addLearningMapSetStatusAliases_(setStatusMap, completedSetIds) {
  /* WM_SET_ID_ALIAS_WM_PREFIX_V1
   * 학습기록_DB Set_ID는 WM4-1-1 형식이고, Map.html set 블록은 4-1-1 형식일 수 있습니다.
   * 두 형식이 모두 같은 진행률을 참조할 수 있도록 API 응답에 별칭을 추가합니다.
   */
  if (!setStatusMap || typeof setStatusMap !== 'object') {
    return;
  }

  Object.keys(setStatusMap).forEach(function(key) {
    var value = setStatusMap[key];
    var text = String(key || '').trim();
    if (!text) return;

    if (/^WM\d+-\d+-\d+$/.test(text)) {
      var noPrefix = text.replace(/^WM/, '');
      if (!setStatusMap[noPrefix]) {
        setStatusMap[noPrefix] = value;
      }
    } else if (/^\d+-\d+-\d+$/.test(text)) {
      var withPrefix = 'WM' + text;
      if (!setStatusMap[withPrefix]) {
        setStatusMap[withPrefix] = value;
      }
    }
  });

  if (Array.isArray(completedSetIds)) {
    var original = completedSetIds.slice();
    original.forEach(function(id) {
      var text = String(id || '').trim();
      if (/^WM\d+-\d+-\d+$/.test(text)) {
        var noPrefix = text.replace(/^WM/, '');
        if (completedSetIds.indexOf(noPrefix) === -1) completedSetIds.push(noPrefix);
      } else if (/^\d+-\d+-\d+$/.test(text)) {
        var withPrefix = 'WM' + text;
        if (completedSetIds.indexOf(withPrefix) === -1) completedSetIds.push(withPrefix);
      }
    });
  }
}


function buildLevelHistoryLevelsForLearningMap_(records, completedSetIds, setStatusMap, studentProfile, officialLevelProgress) {
  /* WM_LEVEL_HISTORY_SERVER_PAYLOAD_20260617_V1
   * 히스토리 보드용 레벨 목록을 서버에서 직접 내려줍니다.
   * Map.html이 현재레벨만 다시 계산해도, 이전 완료 레벨 버튼이 사라지지 않도록
   * records / completedSetIds / setStatusMap / 학생관리_DB 현재세트 / 공식 진행 상태를 모두 병합합니다.
   */
  var levelMap = {};
  var levels = [];

  function addLevel(level) {
    var n = Number(level || 0);
    if (n < 3 || n > 13 || levelMap[n]) return;
    levelMap[n] = true;
    levels.push(n);
  }

  function addLevelFromSetId(setId) {
    var plain = normalizeLevelPlainSetId_(setId);
    if (!plain) return;
    addLevel(Number(plain.split('-')[0] || 0));
  }

  if (Array.isArray(records)) {
    records.forEach(function(record) {
      addLevelFromSetId(record && (record.Set_ID || record.setId || record.set_id));
    });
  }

  if (Array.isArray(completedSetIds)) {
    completedSetIds.forEach(addLevelFromSetId);
  }

  if (setStatusMap && typeof setStatusMap === 'object') {
    Object.keys(setStatusMap).forEach(function(key) {
      var item = setStatusMap[key] || {};
      addLevelFromSetId(item.setId || key);
    });
  }

  addLevelFromSetId(studentProfile && studentProfile['현재세트']);
  addLevelFromSetId(officialLevelProgress && officialLevelProgress['다음공식세트']);
  addLevelFromSetId(officialLevelProgress && officialLevelProgress['다음레벨첫세트']);
  addLevel(officialLevelProgress && officialLevelProgress.level);
  addLevel(officialLevelProgress && officialLevelProgress['다음레벨']);
  addLevel(extractLevelNumberForMap_((studentProfile && studentProfile['학습배정']) || ''));

  /* 현재세트가 5레벨 이상이면 직전 레벨은 이미 열린 이력으로 간주합니다.
     이 방어선이 4→5 복귀 시 4레벨 버튼 소실을 막습니다. */
  var currentSetLevel = 0;
  var currentSetPlain = normalizeLevelPlainSetId_(studentProfile && studentProfile['현재세트']);
  if (currentSetPlain) currentSetLevel = Number(currentSetPlain.split('-')[0] || 0);
  var assignedLevel = extractLevelNumberForMap_((studentProfile && studentProfile['학습배정']) || '');
  var activeLevel = currentSetLevel || assignedLevel || Number(officialLevelProgress && officialLevelProgress.level || 0);
  if (activeLevel > 3) {
    addLevel(activeLevel - 1);
  }

  levels.sort(function(a, b) { return a - b; });
  return levels;
}

function buildStudentProfileForLearningMap_(studentId) {
  var info = getStudentBasicInfoForMap_(studentId);
  var learningMode = info.learningMode || getStudentLearningMode_(studentId, '');

  var profile = {
    학생ID: info.studentId || String(studentId || '').trim(),
    학생이름: info.studentName || '',
    학교: info.school || '',
    학년: info.grade || '',
    Class: info.className || '',
    교사명: info.teacherName || '',
    학습배정: info.learningAssign || '',
    현재세트: normalizeCurrentSetForMap_(info.currentSet, info.learningAssign || ''),
    학습모드: learningMode,
    S4실루엣단계: learningMode.S4실루엣단계 || '',
    S6테스트모드: learningMode.S6테스트모드 || '',
    레벨완료조건: learningMode.레벨완료조건,
    레벨완료횟수: 0,
    순차완주세트: 0,
    순차완주회차: 0,
    레벨회차추가: learningMode.레벨회차추가,
    목표회차: buildLevelTargetRoundGoalFromLearningMode_(learningMode).목표회차
  };

  return wmApplyCurrentProgressCacheToStudentProfile_(profile, wmGetCurrentProgressCacheForStudent_(studentId));
}


function buildRecordObjectByHeaders_(headers, row) {
  var record = {};

  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim();
    if (key) {
      record[key] = String(row[i] || '').trim();
    }
  }

  return record;
}

function buildLearningMapDisplayRecord_(rawRecord) {
  rawRecord = rawRecord || {};

  var status = normalizeLearningStatus_(rawRecord['완료상태']);
  if (!isCompleteStatus(status)) {
    status = '';
  }

  return {
    '학습기록ID': String(rawRecord.__displayRecordId || rawRecord['학습기록ID'] || '').trim(),
    '학습날짜': String(rawRecord['학습날짜'] || '').trim(),
    '학생ID': String(rawRecord['학생ID'] || '').trim(),
    '학생이름': String(rawRecord['학생이름'] || '').trim(),
    '학교': String(rawRecord['학교'] || '').trim(),
    '학년': String(rawRecord['학년'] || '').trim(),
    'Class': String(rawRecord['Class'] || rawRecord['반'] || rawRecord['반명'] || '').trim(),
    '교사명': String(rawRecord['교사명'] || '').trim(),
    'Set_ID': String(rawRecord['Set_ID'] || '').trim(),
    '점수': String(rawRecord['한영주관식'] || rawRecord['점수'] || '').trim(),
    '총소요시간': String(rawRecord['총소요시간'] || '').trim(),
    'Step1_총시간': String(rawRecord['Step1_총시간'] || '').trim(),
    'Step2_총시간': String(rawRecord['Step2_총시간'] || '').trim(),
    'Step3_총시간': String(rawRecord['Step3_총시간'] || '').trim(),
    'Step4_총시간': String(rawRecord['Step4_총시간'] || '').trim(),
    'Step5_총시간': String(rawRecord['Step5_총시간'] || rawRecord['시험전재학습'] || rawRecord['시험전학습'] || '').trim(),
    '시험전재학습': String(rawRecord['시험전재학습'] || rawRecord['Step5_총시간'] || rawRecord['시험전학습'] || '').trim(),
    'Test_총시간': String(rawRecord['Test_총시간'] || '').trim(),
    '영한객관식': String(rawRecord['영한객관식'] || '').trim(),
    '한영객관식': String(rawRecord['한영객관식'] || '').trim(),
    '영한주관식': String(rawRecord['영한주관식'] || '').trim(),
    '한영주관식': String(rawRecord['한영주관식'] || '').trim(),
    '틀린단어': String(rawRecord['틀린단어'] || '').trim(),
    '완료': String(rawRecord['완료'] || '').trim(),
    '완료상태': status,
    '기기정보': normalizeDeviceInfo_(rawRecord['기기정보'] || ''),
    '비고': String(rawRecord['비고'] || '').trim()
  };
}

function isCompleteStatus(status) {
  var text = String(status || '').trim();

  if (!text) {
    return false;
  }

  return (
    text === '완료' ||
    text === '학습완료' ||
    text === 'COMPLETE' ||
    text === 'Completed' ||
    text === 'completed' ||
    text === 'DONE' ||
    text === 'done'
  );
}

function outputResult(e, obj) {
  var callback = '';
  if (e && e.parameter && e.parameter.callback) {
    callback = String(e.parameter.callback).trim();
  }

  var json = JSON.stringify(obj);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}




/* WM_OFFICIAL_ROUND_COMPLETION_COLUMN_20260619_V1
 * 2.학습기록_DB의 '완료' 컬럼을 공식 인정 회차의 단일 기준으로 사용합니다.
 * 저장값: 1회차, 2회차, 3회차, 4회차...
 */
function wmParseOfficialCompleteRound_(value) {
  var text = String(value || '').trim();
  var match = text.match(/(\d+)\s*회차/);
  if (!match) {
    match = text.match(/^(\d+)$/);
  }
  var n = match ? Number(match[1]) : 0;
  return isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function wmFormatOfficialCompleteRound_(round) {
  var n = Number(round || 0);
  if (!isFinite(n) || n < 1) n = 1;
  return Math.floor(n) + '회차';
}

function wmGetOfficialRoundFromRow_(headers, row, completeStatus) {
  var idxCompleteRound = headers.indexOf('완료');
  var explicitRound = idxCompleteRound >= 0 ? wmParseOfficialCompleteRound_(row[idxCompleteRound]) : 0;
  if (explicitRound > 0) {
    return explicitRound;
  }

  /* 기존 기록에는 '완료' 컬럼값이 없으므로, 완료상태=완료인 기존 행은 1회차로만 인정합니다.
   * 이 fallback은 1회차 완료 전 2회차로 잘못 올라가는 문제를 막기 위한 안전장치입니다.
   */
  if (isCompleteStatus(completeStatus)) {
    return 1;
  }

  return 0;
}

function wmBuildOfficialRoundMapFromLearningRecords_(sheet, headers, studentId, level, excludeRecordId, prefetchedValues) {
  var roundMap = {};
  var sequence = getLevelSetSequence_(level);
  var sequenceMap = {};

  for (var s = 0; s < sequence.length; s++) {
    roundMap[sequence[s]] = 0;
    sequenceMap[sequence[s]] = true;
  }

  if (!sheet || !headers || !headers.length || !studentId || !sequence.length) {
    return roundMap;
  }

  var idxRecordId = headers.indexOf('학습기록ID');
  var idxStudentId = headers.indexOf('학생ID');
  var idxSetId = headers.indexOf('Set_ID');
  var idxStatus = headers.indexOf('완료상태');

  if (idxStudentId < 0 || idxSetId < 0 || idxStatus < 0) {
    return roundMap;
  }

  var values = prefetchedValues && prefetchedValues.length ? prefetchedValues : sheet.getDataRange().getDisplayValues();
  var targetStudentId = String(studentId || '').trim().toUpperCase();
  var excludeId = String(excludeRecordId || '').trim();

  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    var rowStudentId = String(row[idxStudentId] || '').trim().toUpperCase();
    if (rowStudentId !== targetStudentId) continue;

    var rowRecordId = idxRecordId >= 0 ? String(row[idxRecordId] || '').trim() : '';
    if (excludeId && rowRecordId === excludeId) continue;

    var plainSetId = normalizeLevelPlainSetId_(row[idxSetId]);
    if (!plainSetId || !sequenceMap[plainSetId]) continue;

    var status = String(row[idxStatus] || '').trim();
    if (!isCompleteStatus(status)) continue;

    var round = wmGetOfficialRoundFromRow_(headers, row, status);
    if (round > Number(roundMap[plainSetId] || 0)) {
      roundMap[plainSetId] = round;
    }
  }

  return roundMap;
}

function wmComputeOfficialCompleteRoundForRecord_(sheet, headers, record, prefetchedValues) {
  record = record || {};

  if (!isCompleteStatus(record['완료상태'])) {
    return '';
  }

  var setId = normalizeLevelPlainSetId_(record['Set_ID']);
  var parsed = parseLearningMapPlainSetId_(setId);
  if (!parsed) {
    return '1회차';
  }

  var sequence = getLevelSetSequence_(parsed.level);
  if (!sequence.length) {
    return '1회차';
  }

  var setIndex = sequence.indexOf(setId);
  if (setIndex < 0) {
    return '1회차';
  }

  var studentId = String(record['학생ID'] || '').trim().toUpperCase();
  var recordId = String(record['학습기록ID'] || '').trim();
  var roundMap = wmBuildOfficialRoundMapFromLearningRecords_(sheet, headers, studentId, parsed.level, recordId, prefetchedValues);
  var currentRound = Number(roundMap[setId] || 0);

  for (var round = 1; round <= 30; round++) {
    if (currentRound >= round) {
      continue;
    }

    var previousRoundComplete = true;
    if (round > 1) {
      for (var a = 0; a < sequence.length; a++) {
        if (Number(roundMap[sequence[a]] || 0) < round - 1) {
          previousRoundComplete = false;
          break;
        }
      }
    }

    if (!previousRoundComplete) {
      continue;
    }

    var previousSetsInThisRoundComplete = true;
    for (var b = 0; b < setIndex; b++) {
      if (Number(roundMap[sequence[b]] || 0) < round) {
        previousSetsInThisRoundComplete = false;
        break;
      }
    }

    if (previousSetsInThisRoundComplete) {
      return wmFormatOfficialCompleteRound_(round);
    }
  }

  return wmFormatOfficialCompleteRound_(currentRound + 1);
}

function wmApplyOfficialCompleteRoundToRecord_(sheet, headers, record, targetRow, prefetchedValues) {
  if (!record || !headers || headers.indexOf('완료') < 0) {
    return record;
  }

  if (!isCompleteStatus(record['완료상태'])) {
    record['완료'] = '';
    return record;
  }

  var idxCompleteRound = headers.indexOf('완료');
  if (targetRow && targetRow > 1 && idxCompleteRound >= 0) {
    var existingRound = String(sheet.getRange(targetRow, idxCompleteRound + 1).getDisplayValue() || '').trim();
    if (existingRound) {
      record['완료'] = existingRound;
      return record;
    }
  }

  record['완료'] = wmComputeOfficialCompleteRoundForRecord_(sheet, headers, record, prefetchedValues);
  return record;
}

/* WM_LEVEL_COMPLETE_CONDITION_COUNT_FINAL_LOCK_20260727_V1
 * 레벨 완료의 최종 판단축은 반드시 아래 두 값을 한 묶음으로 사용합니다.
 * - 레벨완료조건: 학생관리_DB의 현재 목표 회차
 * - 레벨완료횟수: 8.현재진행_DB에 저장된 해당 레벨의 최종 완료 횟수
 * 이전 학습기록 전체 재계산값은 최종 레벨 완료 판단에 사용하지 않습니다.
 */
function wmGetLockedLevelCompletionState_(studentId, level) {
  var state = {
    level: Number(level || 0),
    levelCompleteCount: 0,
    learningRecordId: ''
  };

  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss && ss.getSheetByName('8.현재진행_DB');
    if (!sheet || sheet.getLastRow() < 2) return state;

    var values = sheet.getDataRange().getDisplayValues();
    var headers = values[0].map(function(h) { return String(h || '').trim(); });
    var idxStudent = headers.indexOf('학생ID');
    var idxLevel = headers.indexOf('현재레벨');
    var idxSet = headers.indexOf('기록Set_ID');
    if (idxSet < 0) idxSet = headers.indexOf('Set_ID');
    var idxCount = headers.indexOf('레벨완료횟수');
    var idxRecordId = headers.indexOf('학습기록ID');
    var normalizedStudentId = String(studentId || '').trim().toUpperCase();

    for (var i = 1; i < values.length; i++) {
      var rowStudentId = idxStudent >= 0 ? String(values[i][idxStudent] || '').trim().toUpperCase() : '';
      if (rowStudentId !== normalizedStudentId) continue;

      var storedLevelText = idxLevel >= 0 ? String(values[i][idxLevel] || '').trim() : '';
      var storedSetId = idxSet >= 0 ? normalizeLevelPlainSetId_(values[i][idxSet]) : '';
      var storedLevel = Number((storedLevelText.match(/\d+/) || [])[0] || (storedSetId ? storedSetId.split('-')[0] : 0));
      if (storedLevel !== state.level) return state;

      state.levelCompleteCount = wmNormalizeCurrentProgressRoundValue_(idxCount >= 0 ? values[i][idxCount] : 0);
      state.learningRecordId = idxRecordId >= 0 ? String(values[i][idxRecordId] || '').trim() : '';
      return state;
    }
  } catch (lockedLevelStateErr) {}

  return state;
}

/* WM_OFFICIAL_SEQUENTIAL_TRANSITION_20260714_V1
 * 마지막 세트 TEST 저장 성공 시점에서만 회차/레벨/전체완료를 판정합니다.
 * Map 색상·진행률은 계산하지 않고, 다음 공식 Set_ID 전환만 처리합니다.
 */
function wmBuildOfficialSequentialTransitionAfterTest_(record) {
  record = record || {};
  var result = {
    transitionType: 'NONE',
    nextSetId: '',
    level: 0,
    completedRound: 0,
    targetRounds: 0,
    levelCompleteCondition: '',
    nextLevelCompleteCondition: '',
    conditionCopied: false
  };

  if (!isCompleteStatus(record['완료상태'])) return result;

  var plainSetId = normalizeLevelPlainSetId_(record['Set_ID']);
  var parsed = parseLearningMapPlainSetId_(plainSetId);
  if (!parsed) return result;

  var sequence = getLevelSetSequence_(parsed.level);
  var setIndex = sequence.indexOf(plainSetId);
  if (!sequence.length || setIndex < 0) return result;

  var studentId = String(record['학생ID'] || '').trim().toUpperCase();
  var lockedLevelState = wmGetLockedLevelCompletionState_(studentId, parsed.level);
  var completedRound = wmNormalizeCurrentProgressRoundValue_(lockedLevelState.levelCompleteCount);
  var incomingRecordId = String(record['학습기록ID'] || '').trim();
  var currentCondition = '';
  var nextCondition = '';
  var studentSheet = null;
  var targetRow = -1;
  var idxCurrentCondition = -1;

  try {
    var ss = getLmsSpreadsheet_();
    studentSheet = ss && ss.getSheetByName('1.학생관리_DB');
    if (studentSheet && studentSheet.getLastRow() >= 2) {
      var lastCol = studentSheet.getLastColumn();
      var headers = studentSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
        return String(h || '').trim();
      });
      var idxStudentId = headers.indexOf('학생ID');
      idxCurrentCondition = headers.indexOf('레벨완료조건');
      var idxNextCondition = headers.indexOf('다음레벨완료조건');
      var studentRows = wmGetStudentRowNumbersFast_(studentSheet, headers, studentId);
      targetRow = studentRows.length ? studentRows[0] : -1;

      if (targetRow >= 2) {
        var studentRow = studentSheet.getRange(targetRow, 1, 1, lastCol).getDisplayValues()[0];
        currentCondition = normalizeLevelCompleteCondition_(idxCurrentCondition >= 0 ? studentRow[idxCurrentCondition] : '');
        nextCondition = normalizeLevelCompleteCondition_(idxNextCondition >= 0 ? studentRow[idxNextCondition] : '');
      }
    }
  } catch (studentModeErr) {}

  if (!currentCondition) {
    currentCondition = normalizeLevelCompleteCondition_(getStudentLearningMode_(studentId, '').레벨완료조건);
  }
  if (!nextCondition) nextCondition = currentCondition;

  var targetRounds = getLevelBaseRoundsFromCondition_(currentCondition);
  result.level = parsed.level;
  result.completedRound = completedRound;
  result.targetRounds = targetRounds;
  result.levelCompleteCondition = currentCondition;
  result.nextLevelCompleteCondition = nextCondition;

  if (setIndex < sequence.length - 1) {
    record['__WM_LOCKED_LEVEL_COMPLETE_COUNT'] = completedRound;
    result.transitionType = 'SET_CONTINUE';
    result.nextSetId = sequence[setIndex + 1];
    return result;
  }

  /* 마지막 세트의 신규 완료 기록일 때만 현재진행_DB 완료횟수를 정확히 1회 증가시킵니다.
   * 동일 학습기록ID 재확인은 기존 최종 횟수를 그대로 사용하여 중복 증가를 차단합니다. */
  if (!incomingRecordId || incomingRecordId !== lockedLevelState.learningRecordId) {
    completedRound += 1;
  }
  if (completedRound > targetRounds) completedRound = targetRounds;
  completedRound = wmNormalizeCurrentProgressRoundValue_(completedRound);
  record['__WM_LOCKED_LEVEL_COMPLETE_COUNT'] = completedRound;
  result.completedRound = completedRound;

  if (completedRound < targetRounds) {
    result.transitionType = 'ROUND_CONTINUE';
    result.nextSetId = sequence[0];
    return result;
  }

  if (parsed.level >= 13) {
    result.transitionType = 'COURSE_COMPLETE';
    result.nextSetId = '';
    return result;
  }

  result.transitionType = 'LEVEL_COMPLETE';
  result.nextSetId = getNextLevelFirstSetId_(parsed.level);

  if (studentSheet && targetRow >= 2 && idxCurrentCondition >= 0) {
    studentSheet.getRange(targetRow, idxCurrentCondition + 1).setValue(nextCondition);
    result.conditionCopied = true;
    wmClearRuntimeCachesForStudent_(studentId);
  }

  return result;
}


/* =========================================================
 * WM_STUDY_CHECKPOINT_COMPLETE_FAST_V1_20260819
 * 누적저장과 완료처리를 분리합니다.
 * - Step 저장: 직전 미완료 누적행의 Step시간을 승계한 새 누적 스냅샷 + 8.현재진행_DB 위치 갱신
 * - 세트 완료: 최신 누적행의 Step시간은 보존하고 Test 결과 + 완료 확정 + 다음 공식 Set_ID만 갱신
 * 전체 학습기록 재조회/학생프로필 재조회/현재진행 전체 재계산을 실행하지 않습니다.
 * 기존 saveLearningRecordV2 공식은 삭제하지 않고 비상 복구 기준으로 유지합니다.
 * ========================================================= */
function wmFastFindExactRowByHeader_(sheet, headers, headerName, wantedValue, ignoreCase) {
  if (!sheet || !headers) return -1;
  var idx = headers.indexOf(headerName);
  var lastRow = sheet.getLastRow();
  if (idx < 0 || lastRow < 2) return -1;

  var wanted = String(wantedValue || '').trim();
  if (!wanted) return -1;
  if (ignoreCase) wanted = wanted.toUpperCase();

  var values = sheet.getRange(2, idx + 1, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var current = String(values[i][0] || '').trim();
    if (ignoreCase) current = current.toUpperCase();
    if (current === wanted) return i + 2;
  }
  return -1;
}

function wmFastRowToObject_(headers, row) {
  var obj = {};
  (headers || []).forEach(function(header, i) {
    var key = String(header || '').trim();
    if (key) obj[key] = row && row[i] !== undefined ? row[i] : '';
  });
  return obj;
}

function wmFastSetRowField_(headers, row, headerName, value) {
  var target = wmNormalizeHeaderKeyForCurrentProgress_(headerName);
  var idx = -1;
  for (var i = 0; i < (headers || []).length; i++) {
    if (wmNormalizeHeaderKeyForCurrentProgress_(headers[i]) === target) {
      idx = i;
      break;
    }
  }
  if (idx >= 0) row[idx] = value;
}

function wmFastGenerateLearningRecordId_(sheet, headers, now) {
  var idxRecordId = headers.indexOf('학습기록ID');
  if (idxRecordId < 0) throw new Error('2.학습기록_DB에서 학습기록ID 컬럼을 찾을 수 없습니다.');

  var prefix = 'R' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMM');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '001';

  /* 신규 기록은 항상 2행에 들어가므로 정상 최신 ID가 있으면 전체 DB를 읽지 않고 바로 +1 합니다. */
  var topId = String(sheet.getRange(2, idxRecordId + 1).getDisplayValue() || '').trim();
  if (/^R\d{7}$/.test(topId) && topId.indexOf(prefix) === 0) {
    var seq = Number(topId.slice(5));
    if (isFinite(seq) && seq >= 1 && seq < 999) return prefix + String(seq + 1).padStart(3, '0');
  }

  /* 구 임시ID/월변경 등 예외에서만 기존 안전 공식을 사용합니다. */
  return generateMonthlyLearningRecordId_(sheet, headers, now);
}

function wmFastStudyTimeFields_(payload) {
  return wmBuildOfficialLearningTimeFieldsV2_(
    Array.isArray(payload && payload.completedSteps) ? payload.completedSteps : [],
    payload && payload.officialTimeSummary ? payload.officialTimeSummary : {}
  );
}

function wmFastMergeStudyTimesIntoRecord_(record, timeFields) {
  record = record || {};
  timeFields = timeFields || {};
  var pairs = [
    ['Step1_총시간', 'step1'],
    ['Step2_총시간', 'step2'],
    ['Step3_총시간', 'step3'],
    ['Step4_총시간', 'step4'],
    ['Step5_총시간', 'step5'],
    ['시험전재학습', 'beforeTest'],
    ['시험전학습', 'beforeTest'],
    ['Test_총시간', 'testTime']
  ];
  pairs.forEach(function(pair) {
    var value = String(timeFields[pair[1]] || '').trim();
    if (value) record[pair[0]] = value;
  });

  var totalSeconds = parseTimeTextToSeconds_(record['Step1_총시간'])
    + parseTimeTextToSeconds_(record['Step2_총시간'])
    + parseTimeTextToSeconds_(record['Step3_총시간'])
    + parseTimeTextToSeconds_(record['Step4_총시간'])
    + parseTimeTextToSeconds_(record['Step5_총시간'] || record['시험전재학습'] || record['시험전학습'])
    + parseTimeTextToSeconds_(record['Test_총시간']);
  if (totalSeconds > 0) record['총소요시간'] = formatSecondsForSheet(totalSeconds);
  return record;
}

function wmFastApplyFinalTestOnly_(record, timeFields) {
  record = record || {};
  timeFields = timeFields || {};
  var testTime = String(timeFields.testTime || '').trim();
  if (testTime) record['Test_총시간'] = testTime;

  var totalSeconds = parseTimeTextToSeconds_(record['Step1_총시간'])
    + parseTimeTextToSeconds_(record['Step2_총시간'])
    + parseTimeTextToSeconds_(record['Step3_총시간'])
    + parseTimeTextToSeconds_(record['Step4_총시간'])
    + parseTimeTextToSeconds_(record['Step5_총시간'] || record['시험전재학습'] || record['시험전학습'])
    + parseTimeTextToSeconds_(record['Test_총시간']);
  if (totalSeconds > 0) record['총소요시간'] = formatSecondsForSheet(totalSeconds);
  return record;
}

function wmFastBuildOrUpdateLearningRecord_(sheet, headers, payload, isCompleted) {
  var now = new Date();
  var activeRecordId = String(payload.recordSessionId || payload.recordId || payload['학습기록ID'] || '').trim();
  var resumeSourceRecordId = String(payload.resumeSourceRecordId || payload.previousRecordId || '').trim();
  var studentId = String(payload.studentId || payload.학생ID || '').trim();
  var setId = String(payload.Set_ID || payload.setId || '').trim();

  var activeRow = isOfficialLearningRecordId_(activeRecordId)
    ? wmFastFindExactRowByHeader_(sheet, headers, '학습기록ID', activeRecordId, false)
    : -1;
  var activeValues = activeRow > 1 ? sheet.getRange(activeRow, 1, 1, headers.length).getDisplayValues()[0] : null;
  var activeRecord = activeValues ? wmFastRowToObject_(headers, activeValues) : {};

  if (activeRow > 1) {
    var activeSameStudent = String(activeRecord['학생ID'] || '').trim().toUpperCase() === studentId.toUpperCase();
    var activeSameSet = normalizeStudySetIdForCompare_(activeRecord['Set_ID']) === normalizeStudySetIdForCompare_(setId);
    if (!activeSameStudent || !activeSameSet) throw new Error('현재 학습기록ID의 학생/세트가 일치하지 않습니다.');
    if (!isCompleted && isCompleteStatus(activeRecord['완료상태'])) throw new Error('완료된 학습기록ID에는 중간 Step을 저장할 수 없습니다.');
  }

  var resumeSourceRow = -1;
  var resumeSourceRecord = {};
  if (activeRow < 2 && isOfficialLearningRecordId_(resumeSourceRecordId)) {
    resumeSourceRow = wmFastFindExactRowByHeader_(sheet, headers, '학습기록ID', resumeSourceRecordId, false);
    if (resumeSourceRow > 1) {
      var resumeValues = sheet.getRange(resumeSourceRow, 1, 1, headers.length).getDisplayValues()[0];
      resumeSourceRecord = wmFastRowToObject_(headers, resumeValues);
      var resumeSameStudent = String(resumeSourceRecord['학생ID'] || '').trim().toUpperCase() === studentId.toUpperCase();
      var resumeSameSet = normalizeStudySetIdForCompare_(resumeSourceRecord['Set_ID']) === normalizeStudySetIdForCompare_(setId);
      if (!resumeSameStudent || !resumeSameSet || isCompleteStatus(resumeSourceRecord['완료상태'])) {
        resumeSourceRow = -1;
        resumeSourceRecord = {};
      }
    }
  }

  /* WM_SPLIT_SESSION_CUMULATIVE_ROW_V1_20260825
   * 저장 최소단위는 완료 Step입니다.
   * 같은 Study 접속 안에서는 첫 완료 Step에서 만든 1행을 이후 Step/TEST가 계속 갱신합니다.
   * 학습자가 나갔다가 다시 들어온 경우에는 이전 미완료행을 승계 원본으로만 사용하고,
   * 새 접속에서 최소 1개 Step을 완료했을 때 새 행 1개를 만든 뒤 그 접속 동안 같은 행을 갱신합니다.
   * 미완료 Step 또는 단순 이탈만으로는 새 행을 만들지 않습니다. */
  var targetRow = activeRow > 1 ? activeRow : -1;
  var record = activeRow > 1 ? Object.assign({}, activeRecord) : {};
  var recordId = activeRow > 1 ? activeRecordId : '';
  var needsInsert = activeRow < 2;
  var alreadyCompleted = false;

  if (isCompleted && activeRow > 1) {
    alreadyCompleted = isCompleteStatus(activeRecord['완료상태']) || String(activeRecord['완료'] || '').trim() !== '';
    if (alreadyCompleted) {
      return {
        record:activeRecord,
        recordId:activeRecordId,
        targetRow:activeRow,
        sourceRow:activeRow,
        skippedCompleted:true,
        needsInsert:false
      };
    }
  }

  if (activeRow < 2) {
    if (isCompleted && resumeSourceRow < 2) {
      throw new Error('누적 학습기록 원본 행을 찾을 수 없습니다.');
    }

    recordId = wmFastGenerateLearningRecordId_(sheet, headers, now);
    var fastStudentProfile = resumeSourceRow > 1
      ? {
          studentName:resumeSourceRecord['학생이름'],
          school:resumeSourceRecord['학교'],
          grade:resumeSourceRecord['학년'],
          className:resumeSourceRecord['Class'],
          teacherName:resumeSourceRecord['교사명']
        }
      : (studentId ? (getStudentBasicInfoForMap_(studentId) || {}) : {});

    record = {
      '학습기록ID': recordId,
      '학습날짜': Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      '학생ID': studentId,
      '학생이름': String(payload.studentName || payload.학생이름 || fastStudentProfile.studentName || '').trim(),
      '학교': String(fastStudentProfile.school || payload.school || payload.학교 || '').trim(),
      '학년': String(fastStudentProfile.grade || payload.grade || payload.학년 || '').trim(),
      'Class': String(fastStudentProfile.className || payload.className || payload.Class || payload['반명'] || payload['반'] || '').trim(),
      '교사명': String(fastStudentProfile.teacherName || payload.teacherName || payload.교사명 || '').trim(),
      'Set_ID': setId,
      '점수': '',
      '총소요시간': '',
      'Step1_총시간': '',
      'Step2_총시간': '',
      'Step3_총시간': '',
      'Step4_총시간': '',
      'Step5_총시간': '',
      '시험전재학습': '',
      '시험전학습': '',
      'Test_총시간': '',
      '영한객관식': '',
      '한영객관식': '',
      '영한주관식': '',
      '한영주관식': '',
      '틀린단어': '',
      '완료': '',
      '완료상태': '',
      '기기정보': normalizeDeviceInfo_(payload.userAgent || payload.deviceInfo || ''),
      '비고': ''
    };
  }

  record['학습날짜'] = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  record['학생ID'] = String(record['학생ID'] || studentId || '').trim();
  record['학생이름'] = String(record['학생이름'] || payload.studentName || payload.학생이름 || '').trim();
  record['학교'] = String(record['학교'] || payload.school || payload.학교 || '').trim();
  record['학년'] = String(record['학년'] || payload.grade || payload.학년 || '').trim();
  record['Class'] = String(record['Class'] || payload.className || payload.Class || payload['반명'] || payload['반'] || '').trim();
  record['교사명'] = String(record['교사명'] || payload.teacherName || payload.교사명 || '').trim();
  record['Set_ID'] = String(record['Set_ID'] || setId || '').trim();
  record['기기정보'] = String(record['기기정보'] || normalizeDeviceInfo_(payload.userAgent || payload.deviceInfo || '') || '').trim();

  var fastTimeFields = wmFastStudyTimeFields_(payload);

  if (activeRow < 2 && resumeSourceRow > 1) {
    record = wmMergeExistingLearningRecordTimes_(sheet, headers, resumeSourceRow, record);
    record['학습기록ID'] = recordId;
    record['학습날짜'] = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }

  if (isCompleted) {
    /* WM_FINAL_COMPLETE_ALL_STEP_FIELDS_V1_20260825
     * 완료 payload의 Step1~5 + TEST 공식시간을 같은 누적행에 최종 확정합니다. */
    wmFastMergeStudyTimesIntoRecord_(record, fastTimeFields);

    var scores = Array.isArray(payload.scores) ? payload.scores : [];
    var scoreMap = extractScoreMap(scores);
    record['점수'] = getFinalKoEngWriteScoreFromPayload_(payload, scoreMap);
    record['영한객관식'] = scoreMap.engKoChoice;
    record['한영객관식'] = scoreMap.koEngChoice;
    record['영한주관식'] = scoreMap.engKoWrite;
    record['한영주관식'] = scoreMap.koEngWrite;
    record['틀린단어'] = wmBuildWrongWordsCellValue_(payload);
    record['완료상태'] = '완료';

    return {
      record:record,
      recordId:recordId,
      targetRow:targetRow,
      sourceRow:activeRow > 1 ? activeRow : resumeSourceRow,
      skippedCompleted:false,
      needsInsert:needsInsert
    };
  }

  wmFastMergeStudyTimesIntoRecord_(record, fastTimeFields);

  var writeRow = headers.map(function(header) {
    return record[header] !== undefined ? record[header] : '';
  });

  if (needsInsert) {
    sheet.insertRowBefore(2);
    targetRow = 2;
    sheet.getRange(2, 1, 1, headers.length).setValues([writeRow]);
  } else {
    targetRow = activeRow;
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([writeRow]);
  }

  return {
    record:record,
    recordId:recordId,
    targetRow:targetRow,
    sourceRow:activeRow > 1 ? activeRow : resumeSourceRow,
    skippedCompleted:false,
    needsInsert:needsInsert
  };
}

function wmFastReadCurrentProgress_(ss, studentId) {
  var sheet = ss.getSheetByName('8.현재진행_DB');
  if (!sheet || sheet.getLastColumn() < 1) return {sheet:sheet, headers:[], rowNum:-1, row:[], obj:{}};
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
  var rowNum = wmFastFindExactRowByHeader_(sheet, headers, '학생ID', studentId, true);
  var row = rowNum > 1 ? sheet.getRange(rowNum, 1, 1, headers.length).getDisplayValues()[0] : headers.map(function(){ return ''; });
  return {sheet:sheet, headers:headers, rowNum:rowNum, row:row, obj:wmFastRowToObject_(headers, row)};
}

function wmFastProgressSet_(state, headerName, value) {
  wmFastSetRowField_(state.headers, state.row, headerName, value);
  state.obj[headerName] = value;
}

function wmFastProgressNumber_(obj, key) {
  var n = Number(String(obj && obj[key] !== undefined ? obj[key] : '').replace(/[^0-9.-]/g, ''));
  return isFinite(n) ? n : 0;
}

/* WM_CURRENT_PROGRESS_ROUND_SCORE_V1_20260821
 * 현재진행_DB 점수 공식:
 * - TEST 완료 점수는 해당 회차점수(1/2/3회차점수)와 최근점수에 동시에 기록합니다.
 * - 회차별점수는 같은 레벨 안에서 보존하고, 새 레벨 첫 저장 시 초기화합니다.
 * - 최근점수는 가장 최근 완료 TEST 점수만 유지합니다. */
function wmCurrentProgressRoundScoreHeader_(roundValue) {
  var round = Math.floor(Number(roundValue || 0) || 0);
  return round >= 1 && round <= 3 ? (round + '회차점수') : '';
}

/* WM_LEARNING_MAP_SCORE_JSON_PREBUILT_V2_20260822
 * 학생 Map 진입 중 과거 학습기록을 스캔하지 않습니다.
 * 기존 점수는 관리자 1회 실행 wmRunLearningMapScoreJsonBackfill()로 미리 8.현재진행_DB에 구축하고,
 * 이후 TEST 완료 때는 방금 완료한 Set_ID 1개 점수만 기존 JSON에 즉시 갱신합니다. */
function wmParseLearningMapScoreJson_(value) {
  var state = {v:2, scores:{}};
  var text = String(value || '').trim();
  if (!text) return state;
  try {
    var parsed = JSON.parse(text);
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      state.v = Number(parsed.v || 2) || 2;
      if (parsed.scores && typeof parsed.scores === 'object' && !Array.isArray(parsed.scores)) {
        state.scores = parsed.scores;
      } else {
        state.scores = parsed;
      }
      if (parsed.builtAt) state.builtAt = String(parsed.builtAt || '');
      if (parsed.source) state.source = String(parsed.source || '');
    }
  } catch (err) {}
  return state;
}

function wmSerializeLearningMapScoreJson_(state) {
  state = state || {};
  try {
    return JSON.stringify({
      v:2,
      builtAt:String(state.builtAt || ''),
      source:String(state.source || ''),
      scores:(state.scores && typeof state.scores === 'object' && !Array.isArray(state.scores)) ? state.scores : {}
    });
  } catch (err) {
    return '{"v":2,"builtAt":"","source":"","scores":{}}';
  }
}

function wmUpdateLearningMapScoreJson_(existingJson, setId, roundValue, scoreValue) {
  var state = wmParseLearningMapScoreJson_(existingJson);
  var plainSetId = normalizeLevelPlainSetId_(setId);
  var round = Math.max(1, Math.min(30, Math.floor(Number(roundValue || 1) || 1)));
  var score = wmNormalizeOfficialScoreValue_(scoreValue);
  if (!plainSetId || !score) return wmSerializeLearningMapScoreJson_(state);

  var entry = state.scores[plainSetId];
  if (!entry || typeof entry !== 'object' || Array.isArray(entry)) entry = {};
  entry[String(round)] = score;
  entry.latest = score;
  state.scores[plainSetId] = entry;
  state.v = 2;
  return wmSerializeLearningMapScoreJson_(state);
}

function wmRunLearningMapScoreJsonBackfill() {
  /* 관리자 1회 사전복구 전용입니다.
   * 학생 로그인/Map/Study 실행경로에서는 호출하지 않습니다.
   * 2.학습기록_DB 전체를 한 번 읽어 학생별 Set_ID 점수지도를 만든 뒤
   * 8.현재진행_DB 학습맵점수JSON 열에 일괄 저장합니다. */
  var ss = getLmsSpreadsheet_();
  if (!ss) throw new Error('LMS 스프레드시트 연결 실패');

  var recordSheet = ss.getSheetByName('2.학습기록_DB');
  var progressSheet = ss.getSheetByName('8.현재진행_DB');
  if (!recordSheet) throw new Error('2.학습기록_DB 시트를 찾을 수 없습니다.');
  if (!progressSheet) throw new Error('8.현재진행_DB 시트를 찾을 수 없습니다.');

  var recordValues = recordSheet.getDataRange().getDisplayValues();
  if (!recordValues || !recordValues.length) throw new Error('2.학습기록_DB 데이터가 없습니다.');
  var recordHeaders = recordValues[0].map(function(h){ return String(h || '').trim(); });

  var idxStudent = recordHeaders.indexOf('학생ID');
  var idxSet = recordHeaders.indexOf('Set_ID');
  var idxScore = recordHeaders.indexOf('점수');
  var idxKoEngWrite = recordHeaders.indexOf('한영주관식');
  var idxRound = recordHeaders.indexOf('완료');
  var idxStatus = recordHeaders.indexOf('완료상태');
  if (idxStudent < 0 || idxSet < 0) throw new Error('학습기록_DB 학생ID/Set_ID 헤더를 확인하세요.');

  var statesByStudent = {};
  var scoreCount = 0;

  /* 학습기록_DB는 최신행이 위쪽이므로 같은 세트/회차는 첫 점수를 우선 보존합니다. */
  for (var r = 1; r < recordValues.length; r++) {
    var row = recordValues[r] || [];
    var studentId = String(row[idxStudent] || '').trim().toUpperCase();
    var setId = normalizeLevelPlainSetId_(row[idxSet]);
    if (!studentId || !setId) continue;

    var score = wmNormalizeOfficialScoreValue_(idxScore >= 0 ? row[idxScore] : '') ||
      wmNormalizeOfficialScoreValue_(idxKoEngWrite >= 0 ? row[idxKoEngWrite] : '');
    if (!score) continue;

    var state = statesByStudent[studentId];
    if (!state) {
      state = statesByStudent[studentId] = {v:2, builtAt:'', source:'ADMIN_BACKFILL', scores:{}};
    }

    var entry = state.scores[setId];
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) entry = {};

    if (!entry.latest) entry.latest = score;

    var round = idxRound >= 0 ? wmParseOfficialCompleteRound_(row[idxRound]) : 0;
    var status = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : '';
    if (!round && isCompleteStatus(status)) round = 1;
    if (round > 0 && !entry[String(round)]) entry[String(round)] = score;

    state.scores[setId] = entry;
    scoreCount += 1;
  }

  var progressValues = progressSheet.getDataRange().getDisplayValues();
  if (!progressValues || !progressValues.length) throw new Error('8.현재진행_DB 데이터가 없습니다.');
  var progressHeaders = progressValues[0].map(function(h){ return String(h || '').trim(); });
  var idxProgressStudent = progressHeaders.indexOf('학생ID');
  if (idxProgressStudent < 0) throw new Error('8.현재진행_DB 학생ID 헤더를 확인하세요.');

  var idxScoreJson = progressHeaders.indexOf('학습맵점수JSON');
  if (idxScoreJson < 0) {
    idxScoreJson = progressHeaders.length;
    progressSheet.getRange(1, idxScoreJson + 1).setValue('학습맵점수JSON');
    progressHeaders.push('학습맵점수JSON');
  }

  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var output = [];
  var studentCount = 0;
  var setCount = 0;
  var cacheKeys = [wmCacheKey_('CP_SET_MAP', 'ALL')];

  for (var p = 1; p < progressValues.length; p++) {
    var progressStudentId = String(progressValues[p][idxProgressStudent] || '').trim().toUpperCase();
    var stateForStudent = statesByStudent[progressStudentId] || {v:2, scores:{}};
    stateForStudent.v = 2;
    stateForStudent.builtAt = nowText;
    stateForStudent.source = 'ADMIN_BACKFILL';

    var setKeys = Object.keys(stateForStudent.scores || {});
    if (progressStudentId) {
      studentCount += 1;
      setCount += setKeys.length;
      cacheKeys.push(wmCacheKey_('CURRENT_PROGRESS_STUDENT', progressStudentId));
    }
    output.push([wmSerializeLearningMapScoreJson_(stateForStudent)]);
  }

  if (output.length) {
    progressSheet.getRange(2, idxScoreJson + 1, output.length, 1).setValues(output);
  }

  var cache = wmGetScriptCache_();
  if (cache && cacheKeys.length) {
    for (var c = 0; c < cacheKeys.length; c += 90) {
      try { cache.removeAll(cacheKeys.slice(c, c + 90)); } catch (cacheErr) {}
    }
  }

  return {
    success:true,
    header:'학습맵점수JSON',
    currentProgressStudents:studentCount,
    restoredSets:setCount,
    scoreRowsRead:scoreCount,
    runtimeBackfill:false,
    message:'학습맵 점수 사전복구 완료'
  };
}

function wmClearCurrentProgressRoundScores_(state) {
  ['1회차점수','2회차점수','3회차점수'].forEach(function(headerName) {
    wmFastProgressSet_(state, headerName, '');
  });
}

function wmNormalizeCurrentProgressRecordSetId_(candidateSetId, fallbackSetId) {
  var candidate = String(candidateSetId || '').trim();
  var fallback = String(fallbackSetId || '').trim();
  var plain = normalizeLevelPlainSetId_(candidate);

  if (!plain || !isValidLearningMapSetId_(plain)) {
    candidate = fallback;
    plain = normalizeLevelPlainSetId_(candidate);
  }
  if (!plain || !isValidLearningMapSetId_(plain)) return '';

  var useWmPrefix = /^WM/i.test(String(candidateSetId || '').trim()) || /^WM/i.test(fallback);
  return useWmPrefix ? ('WM' + plain) : plain;
}

function wmResolveCurrentProgressStudentInfo_(studentId, payload, record, currentObj) {
  payload = payload || {};
  record = record || {};
  currentObj = currentObj || {};

  var info = {
    studentName: String(currentObj['학생이름'] || record['학생이름'] || payload.studentName || payload.학생이름 || '').trim(),
    school: String(currentObj['학교'] || record['학교'] || payload.school || payload.학교 || '').trim(),
    grade: String(currentObj['학년'] || record['학년'] || payload.grade || payload.학년 || '').trim(),
    className: String(currentObj['Class'] || record['Class'] || payload.className || payload.Class || payload['반명'] || payload['반'] || '').trim(),
    teacherName: String(currentObj['교사명'] || record['교사명'] || payload.teacherName || payload.교사명 || '').trim()
  };

  var needsProfile = !info.studentName || (!info.school && !info.grade && !info.className && !info.teacherName);
  if (needsProfile && studentId) {
    try {
      var profile = getStudentBasicInfoForMap_(studentId) || {};
      info.studentName = info.studentName || String(profile.studentName || '').trim();
      info.school = info.school || String(profile.school || '').trim();
      info.grade = info.grade || String(profile.grade || '').trim();
      info.className = info.className || String(profile.className || '').trim();
      info.teacherName = info.teacherName || String(profile.teacherName || '').trim();
    } catch (ignoreProfileErr) {}
  }

  return info;
}

function wmFastStudyStepPercent_(completedStep) {
  var step = wmNormalizeCurrentProgressCompletedStep_(completedStep);
  if (step === 'STEP1') return 17;
  if (step === 'STEP2') return 33;
  if (step === 'STEP3') return 50;
  if (step === 'STEP4') return 67;
  if (step === 'STEP5') return 83;
  if (step === 'TEST') return 100;
  return 0;
}

function wmFastCopyNextLevelCondition_(studentId) {
  var result = {conditionCopied:false, nextCondition:''};
  try {
    var ss = getLmsSpreadsheet_();
    var sheet = ss && ss.getSheetByName('1.학생관리_DB');
    if (!sheet || sheet.getLastRow() < 2) return result;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
    var rowNum = wmFastFindExactRowByHeader_(sheet, headers, '학생ID', studentId, true);
    if (rowNum < 2) return result;
    var idxCurrent = headers.indexOf('레벨완료조건');
    var idxNext = headers.indexOf('다음레벨완료조건');
    if (idxCurrent < 0) return result;
    var row = sheet.getRange(rowNum, 1, 1, headers.length).getDisplayValues()[0];
    var currentCondition = normalizeLevelCompleteCondition_(idxCurrent >= 0 ? row[idxCurrent] : '');
    var nextCondition = normalizeLevelCompleteCondition_(idxNext >= 0 ? row[idxNext] : '');
    if (!nextCondition) nextCondition = currentCondition;
    if (nextCondition) {
      sheet.getRange(rowNum, idxCurrent + 1).setValue(nextCondition);
      result.conditionCopied = true;
      result.nextCondition = nextCondition;
    }
  } catch (err) {}
  return result;
}

function wmFastResolveCompletionTransition_(payload, progressObj, setId) {
  var plainSetId = normalizeLevelPlainSetId_(setId);
  var parsed = parseLearningMapPlainSetId_(plainSetId);
  var result = {
    transitionType:'SET_CONTINUE', nextSetId:'', level:parsed ? parsed.level : 0,
    completedRound:0, recordRound:1, targetRounds:Math.max(1, Number(payload.목표회차 || payload.targetRounds || 1) || 1),
    conditionCopied:false, nextLevelCompleteCondition:''
  };
  if (!parsed) return result;

  var sequence = getLevelSetSequence_(parsed.level) || [];
  var setIndex = sequence.indexOf(plainSetId);
  var completedRounds = wmNormalizeCurrentProgressRoundValue_(wmFastProgressNumber_(progressObj, '레벨완료횟수'));
  var targetRounds = Math.max(1, Number(payload.목표회차 || payload.targetRounds || 0) || getLevelBaseRoundsFromCondition_(progressObj['레벨완료조건'] || payload.레벨완료조건 || '2회'));
  result.targetRounds = targetRounds;
  result.recordRound = Math.min(targetRounds, completedRounds + 1);
  result.completedRound = completedRounds;

  if (!sequence.length || setIndex < 0) {
    result.nextSetId = String(payload.nextSetId || '').trim();
    return result;
  }

  if (setIndex < sequence.length - 1) {
    result.transitionType = 'SET_CONTINUE';
    result.nextSetId = sequence[setIndex + 1];
    return result;
  }

  var newCompletedRounds = Math.min(targetRounds, completedRounds + 1);
  result.completedRound = newCompletedRounds;
  result.recordRound = newCompletedRounds;
  if (newCompletedRounds < targetRounds) {
    result.transitionType = 'ROUND_CONTINUE';
    result.nextSetId = sequence[0];
    return result;
  }

  if (parsed.level >= 13) {
    result.transitionType = 'COURSE_COMPLETE';
    result.nextSetId = '';
    return result;
  }

  result.transitionType = 'LEVEL_COMPLETE';
  result.nextSetId = getNextLevelFirstSetId_(parsed.level);
  var copied = wmFastCopyNextLevelCondition_(String(payload.studentId || payload.학생ID || '').trim());
  result.conditionCopied = !!copied.conditionCopied;
  result.nextLevelCompleteCondition = copied.nextCondition || '';
  return result;
}

function wmFastUpdateHistoryForCompletedLevel_(progressObj, completedLevel, completedRounds) {
  var rawLevels = String(progressObj['히스토리레벨'] || progressObj['완료된레벨'] || '').split('|');
  var rawCounts = String(progressObj['히스토리횟수'] || progressObj['완료횟수'] || '').split('|');
  var levels = [];
  var counts = [];
  rawLevels.forEach(function(value, index) {
    var levelText = String(value || '').trim();
    if (!/^(3|4|5|6|7|8|9|10|11|12|13)$/.test(levelText)) return;
    var existingIndex = levels.indexOf(levelText);
    if (existingIndex < 0) {
      levels.push(levelText);
      counts.push(String(wmNormalizeCurrentProgressRoundValue_(rawCounts[index])));
    }
  });
  var levelText = String(completedLevel || '');
  if (!/^(3|4|5|6|7|8|9|10|11|12|13)$/.test(levelText)) return {levels:levels.join('|'), counts:counts.join('|')};
  var idx = levels.indexOf(levelText);
  var countText = String(wmNormalizeCurrentProgressRoundValue_(completedRounds));
  if (idx < 0) {
    levels.push(levelText);
    counts.push(countText);
  } else {
    counts[idx] = countText;
  }
  return {levels:levels.join('|'), counts:counts.join('|')};
}

function wmFastWriteCurrentProgress_(ss, payload, recordId, isCompleted, transition, finalRecord) {
  var studentId = String(payload.studentId || payload.학생ID || '').trim();
  var setId = wmNormalizeCurrentProgressRecordSetId_(payload.Set_ID || payload.setId || '', '');
  var state = wmFastReadCurrentProgress_(ss, studentId);
  if (!state.sheet || !state.headers.length) return {success:false, message:'8.현재진행_DB 없음'};
  if (!setId) return {success:false, message:'현재진행_DB Set_ID 확인 실패'};

  /* WM_CURRENT_PROGRESS_ORDER_TIME_MS_FAST_V1_20260821
   * 실제 학습 저장 경로도 최종수정일을 밀리초까지 기록합니다. */
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  var lastCompletedStep = wmNormalizeCurrentProgressCompletedStep_(payload.lastCompletedStep || '');
  var parsed = parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(setId));
  var sequence = parsed ? (getLevelSetSequence_(parsed.level) || []) : [];
  var totalSetCount = sequence.length || wmFastProgressNumber_(state.obj, '전체세트수');

  /* WM_CURRENT_PROGRESS_FAST_PROFILE_MERGE_20260820_V1
   * 빠른 저장경로에서도 현재진행_DB 학생정보 누락을 막습니다.
   * 이미 만들어진 학습기록/현재진행 값을 우선 재사용하고, 핵심정보가 모두 비었을 때만 학생관리_DB를 조회합니다. */
  var progressStudentInfo = wmResolveCurrentProgressStudentInfo_(studentId, payload, finalRecord, state.obj);
  wmFastProgressSet_(state, '학생ID', String(state.obj['학생ID'] || studentId).trim());
  if (!String(state.obj['학생이름'] || '').trim()) wmFastProgressSet_(state, '학생이름', progressStudentInfo.studentName);
  if (!String(state.obj['학교'] || '').trim()) wmFastProgressSet_(state, '학교', progressStudentInfo.school);
  if (!String(state.obj['학년'] || '').trim()) wmFastProgressSet_(state, '학년', progressStudentInfo.grade);
  if (!String(state.obj['Class'] || '').trim()) wmFastProgressSet_(state, 'Class', progressStudentInfo.className);
  if (!String(state.obj['교사명'] || '').trim()) wmFastProgressSet_(state, '교사명', progressStudentInfo.teacherName);
  if (!String(state.obj['레벨완료조건'] || '').trim()) wmFastProgressSet_(state, '레벨완료조건', String(payload.레벨완료조건 || '').trim());
  if (!String(state.obj['S4실루엣단계'] || '').trim()) wmFastProgressSet_(state, 'S4실루엣단계', String(payload.S4실루엣단계 || '').trim());
  if (!String(state.obj['S6테스트모드'] || '').trim()) wmFastProgressSet_(state, 'S6테스트모드', String(payload.S6테스트모드 || '').trim());

  var previousProgressSet = parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(state.obj['Set_ID'] || ''));
  if (previousProgressSet && parsed && previousProgressSet.level !== parsed.level) {
    wmClearCurrentProgressRoundScores_(state);
  }

  wmFastProgressSet_(state, 'Set_ID', setId);
  wmFastProgressSet_(state, '학습기록ID', recordId);
  wmFastProgressSet_(state, '최종수정일', nowText);
  wmFastProgressSet_(state, '학습날짜', nowText);

  if (!isCompleted) {
    var currentStep = wmGetCurrentStepAfterCompletedStep_(lastCompletedStep);
    wmFastProgressSet_(state, '완료Step', lastCompletedStep);
    wmFastProgressSet_(state, '현재Step', currentStep);
    wmFastProgressSet_(state, '기록Set_ID', wmNormalizeCurrentProgressRecordSetId_(setId, setId));
    wmFastProgressSet_(state, '현재세트', normalizeLevelPlainSetId_(setId));
    wmFastProgressSet_(state, '세트진행률', wmFormatCurrentProgressPercentText_(wmFastStudyStepPercent_(lastCompletedStep)));
  } else {
    transition = transition || {};
    var nextSetId = String(transition.nextSetId || payload.nextSetId || '').trim();
    var transitionType = String(transition.transitionType || 'SET_CONTINUE').toUpperCase();
    var currentLevel = parsed ? parsed.level : 0;

    /* WM_CURRENT_PROGRESS_COMPLETION_NEXTSET_GUARD_20260820_V1
     * TEST 완료의 SET_CONTINUE는 payload/재시도 값이 아니라
     * 방금 완료한 Set_ID의 공식 순서에서 다음 세트를 다시 확정합니다.
     * 예: 3-1-10 -> 3-2-1. 3-1-2 같은 역행값은 현재진행_DB에 기록하지 않습니다. */
    if (transitionType === 'SET_CONTINUE' && parsed && sequence.length) {
      var completedSetIndex = sequence.indexOf(normalizeLevelPlainSetId_(setId));
      if (completedSetIndex >= 0 && completedSetIndex < sequence.length - 1) {
        nextSetId = sequence[completedSetIndex + 1];
      }
    }

    var nextParsed = nextSetId ? parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(nextSetId)) : null;
    var existingCompletedSetCount = wmFastProgressNumber_(state.obj, '완료세트수');
    var existingSequentialSet = wmFastProgressNumber_(state.obj, '순차완주세트');
    var existingSequentialRound = wmFastProgressNumber_(state.obj, '순차완주회차');
    var completedSetCount = totalSetCount ? Math.min(totalSetCount, existingCompletedSetCount + 1) : existingCompletedSetCount + 1;
    var levelProgress = totalSetCount > 0 ? Math.round((completedSetCount / totalSetCount) * 100) : 0;

    wmFastProgressSet_(state, '완료Step', 'TEST');
    wmFastProgressSet_(state, '현재Step', 'STEP1');
    wmFastProgressSet_(state, '기록Set_ID', wmNormalizeCurrentProgressRecordSetId_(nextSetId || setId, setId));
    var completedRoundForScore = Number(transition.recordRound || 1);
    var completedRoundScore = String(finalRecord && finalRecord['점수'] || '').trim();
    var completedRoundScoreHeader = wmCurrentProgressRoundScoreHeader_(completedRoundForScore);
    wmFastProgressSet_(state, '완료회차', completedRoundForScore);
    if (completedRoundScore && completedRoundScoreHeader) {
      wmFastProgressSet_(state, completedRoundScoreHeader, completedRoundScore);
      wmFastProgressSet_(state, '최근점수', completedRoundScore);
      if (state.headers.indexOf('학습맵점수JSON') >= 0) {
        wmFastProgressSet_(
          state,
          '학습맵점수JSON',
          wmUpdateLearningMapScoreJson_(
            state.obj['학습맵점수JSON'] || '',
            setId,
            completedRoundForScore,
            completedRoundScore
          )
        );
      }
    }
    wmFastProgressSet_(state, '세트진행률', wmFormatCurrentProgressPercentText_(0));
    wmFastProgressSet_(state, '다음세트', nextSetId);

    if (transitionType === 'LEVEL_COMPLETE') {
      var history = wmFastUpdateHistoryForCompletedLevel_(state.obj, currentLevel, transition.completedRound || transition.recordRound || 1);
      var nextLevel = nextParsed ? nextParsed.level : currentLevel + 1;
      var nextTotal = getLevelSetSequence_(nextLevel).length;
      wmFastProgressSet_(state, '현재레벨', nextLevel + '레벨');
      wmFastProgressSet_(state, '현재세트', nextSetId ? normalizeLevelPlainSetId_(nextSetId) : '');
      wmFastProgressSet_(state, '레벨완료횟수', '0');
      wmFastProgressSet_(state, '순차완주세트', 0);
      wmFastProgressSet_(state, '순차완주회차', '0');
      wmFastProgressSet_(state, '완료세트수', 0);
      wmFastProgressSet_(state, '전체세트수', nextTotal);
      wmFastProgressSet_(state, '레벨진행률', wmFormatCurrentProgressPercentText_(0));
      wmFastProgressSet_(state, '레벨완료여부', 'N');
      if (transition.nextLevelCompleteCondition) wmFastProgressSet_(state, '레벨완료조건', transition.nextLevelCompleteCondition);
      wmFastProgressSet_(state, '히스토리레벨', history.levels);
      wmFastProgressSet_(state, '히스토리횟수', history.counts);
      wmFastProgressSet_(state, '완료된레벨', history.levels);
      wmFastProgressSet_(state, '완료횟수', history.counts);
    } else if (transitionType === 'COURSE_COMPLETE') {
      var courseHistory = wmFastUpdateHistoryForCompletedLevel_(state.obj, currentLevel, transition.completedRound || transition.recordRound || 1);
      wmFastProgressSet_(state, '현재레벨', currentLevel + '레벨');
      wmFastProgressSet_(state, '현재세트', normalizeLevelPlainSetId_(setId));
      wmFastProgressSet_(state, '레벨완료횟수', String(transition.completedRound || transition.recordRound || 1));
      wmFastProgressSet_(state, '순차완주세트', totalSetCount);
      wmFastProgressSet_(state, '순차완주회차', String(transition.completedRound || transition.recordRound || 1));
      wmFastProgressSet_(state, '완료세트수', totalSetCount);
      wmFastProgressSet_(state, '전체세트수', totalSetCount);
      wmFastProgressSet_(state, '레벨진행률', wmFormatCurrentProgressPercentText_(100));
      wmFastProgressSet_(state, '레벨완료여부', 'Y');
      wmFastProgressSet_(state, '히스토리레벨', courseHistory.levels);
      wmFastProgressSet_(state, '히스토리횟수', courseHistory.counts);
      wmFastProgressSet_(state, '완료된레벨', courseHistory.levels);
      wmFastProgressSet_(state, '완료횟수', courseHistory.counts);
    } else if (transitionType === 'ROUND_CONTINUE') {
      wmFastProgressSet_(state, '현재레벨', currentLevel + '레벨');
      wmFastProgressSet_(state, '현재세트', normalizeLevelPlainSetId_(nextSetId || setId));
      wmFastProgressSet_(state, '레벨완료횟수', String(transition.completedRound || 0));
      wmFastProgressSet_(state, '순차완주세트', 0);
      wmFastProgressSet_(state, '순차완주회차', String(Math.min(transition.targetRounds || 1, (transition.completedRound || 0) + 1)));
      wmFastProgressSet_(state, '완료세트수', completedSetCount);
      wmFastProgressSet_(state, '전체세트수', totalSetCount);
      wmFastProgressSet_(state, '레벨진행률', wmFormatCurrentProgressPercentText_(levelProgress));
      wmFastProgressSet_(state, '레벨완료여부', 'N');
    } else {
      wmFastProgressSet_(state, '현재레벨', currentLevel ? (currentLevel + '레벨') : String(state.obj['현재레벨'] || ''));
      wmFastProgressSet_(state, '현재세트', normalizeLevelPlainSetId_(nextSetId || setId));
      wmFastProgressSet_(state, '레벨완료횟수', String(wmFastProgressNumber_(state.obj, '레벨완료횟수')));
      wmFastProgressSet_(state, '순차완주세트', totalSetCount ? Math.min(totalSetCount, existingSequentialSet + 1) : existingSequentialSet + 1);
      wmFastProgressSet_(state, '순차완주회차', String(existingSequentialRound || Math.max(1, wmFastProgressNumber_(state.obj, '레벨완료횟수') + 1)));
      wmFastProgressSet_(state, '완료세트수', completedSetCount);
      wmFastProgressSet_(state, '전체세트수', totalSetCount);
      wmFastProgressSet_(state, '레벨진행률', wmFormatCurrentProgressPercentText_(levelProgress));
      wmFastProgressSet_(state, '레벨완료여부', 'N');
    }
  }

  /* WM_HISTORY_SAVE_CURRENT_LEVEL_V1
   * 기존 한 번의 현재진행_DB 저장에 히스토리 두 항목을 함께 반영합니다.
   * 과거 레벨/횟수는 보존하고, 새 레벨 미학습은 0, 학습 이후는 실제 진행회차입니다.
   * 추가 DB 조회/저장 또는 학습기록 재계산은 하지 않습니다. */
  var historyCurrentLevel = Number(String(state.obj['현재레벨'] || '').replace(/[^0-9]/g, ''));
  if (historyCurrentLevel >= 3 && historyCurrentLevel <= 13) {
    var historyCurrentRound = wmNormalizeCurrentProgressRoundValue_(state.obj['순차완주회차']);
    if (!(isCompleted && transitionType === 'LEVEL_COMPLETE')) {
      historyCurrentRound = Math.max(1, historyCurrentRound);
    }
    var currentHistory = wmFastUpdateHistoryForCompletedLevel_(state.obj, historyCurrentLevel, historyCurrentRound);
    wmFastProgressSet_(state, '히스토리레벨', currentHistory.levels);
    wmFastProgressSet_(state, '히스토리횟수', currentHistory.counts);
  }

  if (state.rowNum > 1) {
    state.sheet.getRange(state.rowNum, 1, 1, state.headers.length).setValues([state.row]);
  } else {
    state.sheet.insertRowBefore(2);
    state.rowNum = 2;
    state.sheet.getRange(2, 1, 1, state.headers.length).setValues([state.row]);
  }

  return {
    success:true,
    rowNum:state.rowNum,
    nextSetId:String(state.obj['기록Set_ID'] || '').trim(),
    currentProgress:Object.assign({}, state.obj),
    currentProgressCache:Object.assign({}, state.obj),
    currentProgressRow:Object.assign({}, state.obj)
  };
}

function wmFastIsCurrentProgressCompletionCommitted_(progressObj, recordSetId) {
  progressObj = progressObj || {};
  var savedSetId = normalizeLevelPlainSetId_(progressObj['Set_ID'] || '');
  var targetSetId = normalizeLevelPlainSetId_(recordSetId || '');
  var completedStep = wmNormalizeCurrentProgressCompletedStep_(progressObj['완료Step'] || '');
  var currentStep = String(progressObj['현재Step'] || '').trim().toUpperCase();
  var recordStepSetId = normalizeLevelPlainSetId_(wmNormalizeCurrentProgressRecordSetId_(progressObj['기록Set_ID'] || '', ''));
  if (!targetSetId || savedSetId !== targetSetId || completedStep !== 'TEST' || currentStep !== 'STEP1' || !recordStepSetId) return false;

  var parsed = parseLearningMapPlainSetId_(targetSetId);
  var sequence = parsed ? (getLevelSetSequence_(parsed.level) || []) : [];
  var index = sequence.indexOf(targetSetId);
  if (!parsed || index < 0 || !sequence.length) return false;
  if (index < sequence.length - 1) return recordStepSetId === sequence[index + 1];
  if (parsed.level >= 13) return recordStepSetId === targetSetId;
  return recordStepSetId === sequence[0] || recordStepSetId === normalizeLevelPlainSetId_(getNextLevelFirstSetId_(parsed.level));
}

function wmFastInferCompletedRetryResponse_(progressObj, record, payload) {
  var setId = String(record['Set_ID'] || payload.Set_ID || payload.setId || '').trim();
  var nextSetId = String(progressObj['기록Set_ID'] || progressObj['다음세트'] || payload.nextSetId || '').trim();
  var currentParsed = parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(setId));
  var nextParsed = nextSetId ? parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(nextSetId)) : null;
  var transitionType = 'SET_CONTINUE';
  if (!nextSetId && currentParsed && currentParsed.level >= 13) transitionType = 'COURSE_COMPLETE';
  else if (currentParsed && nextParsed && nextParsed.level > currentParsed.level) transitionType = 'LEVEL_COMPLETE';
  else if (currentParsed && nextParsed && nextParsed.level === currentParsed.level) {
    var seq = getLevelSetSequence_(currentParsed.level) || [];
    if (seq.length && normalizeLevelPlainSetId_(setId) === seq[seq.length - 1] && normalizeLevelPlainSetId_(nextSetId) === seq[0]) transitionType = 'ROUND_CONTINUE';
  }
  return {
    success:true, alreadyCompleted:true, recordId:String(record['학습기록ID'] || '').trim(), setId:setId,
    status:'완료', transitionType:transitionType, nextOfficialSetId:nextSetId,
    completedRound:wmParseOfficialCompleteRound_(record['완료']) || wmFastProgressNumber_(progressObj, '완료회차') || 1,
    targetRounds:Math.max(1, Number(payload.목표회차 || payload.targetRounds || 1) || 1),
    transitionLevel:currentParsed ? currentParsed.level : 0
  };
}

function wmSaveStudyCheckpointFast_(e) {
  try {
    var payload = readLearningPayload(e);
    if (!payload) return outputResult(e, {success:false, message:'학습 완료 payload가 없습니다.'});

    var studentId = String(payload.studentId || payload.학생ID || '').trim();
    var sessionToken = String(payload.sessionToken || payload.현재세션 || payload.wmSessionToken || '').trim();
    if (studentId && sessionToken) {
      var sessionState = getStudentSessionState_(studentId);
      if (sessionState.success) {
        var valid = sessionState.currentSession === sessionToken && String(sessionState.sessionStatus || '').toUpperCase() === 'LOGIN';
        if (!valid) return outputResult(e, wmBuildSessionExpiredResponse_());
      }
    }

    var ss = getLmsSpreadsheet_();
    if (!ss) return outputResult(e, {success:false, message:'스프레드시트 연결 실패'});
    var recordSheet = ss.getSheetByName('2.학습기록_DB');
    if (!recordSheet) return outputResult(e, {success:false, message:'2.학습기록_DB 없음'});
    var recordHeaders = recordSheet.getRange(1, 1, 1, recordSheet.getLastColumn()).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
    recordHeaders = normalizeLearningRecordHeaders_(recordHeaders);

    var isCompleted = normalizeLearningStatus_(payload.status) === '완료';
    var progressBefore = wmFastReadCurrentProgress_(ss, studentId);
    var completionSetIdBefore = String(payload.Set_ID || payload.setId || '').trim();

    /* TEST 완료 재요청은 현재진행_DB가 이미 완료 확정된 경우 새 행을 다시 만들지 않습니다. */
    if (isCompleted && wmFastIsCurrentProgressCompletionCommitted_(progressBefore.obj || {}, completionSetIdBefore)) {
      var committedRecordId = String(progressBefore.obj && progressBefore.obj['학습기록ID'] || '').trim();
      var committedRow = committedRecordId
        ? wmFastFindExactRowByHeader_(recordSheet, recordHeaders, '학습기록ID', committedRecordId, false)
        : -1;
      if (committedRow > 1) {
        var committedValues = recordSheet.getRange(committedRow, 1, 1, recordHeaders.length).getDisplayValues()[0];
        var committedRecord = wmFastRowToObject_(recordHeaders, committedValues);
        var committedResponse = wmFastInferCompletedRetryResponse_(progressBefore.obj || {}, committedRecord, payload);
        committedResponse.currentProgress = Object.assign({}, progressBefore.obj || {});
        committedResponse.currentProgressCache = Object.assign({}, progressBefore.obj || {});
        committedResponse.currentProgressRow = Object.assign({}, progressBefore.obj || {});
        committedResponse.learningRecord = Object.assign({}, committedRecord || {});
        wmClearRuntimeCachesForStudent_(studentId);
        return outputResult(e, committedResponse);
      }
    }

    if (isCompleted &&
        !String(payload.recordSessionId || '').trim() &&
        !String(payload.resumeSourceRecordId || '').trim()) {
      payload.recordSessionId = String(progressBefore.obj && progressBefore.obj['학습기록ID'] || '').trim();
    }
    var recordResult = WM_STUDY_RUNTIME_EXECUTION_SERVER_LOCK_V1.SAVE.fastBuildOrUpdate(recordSheet, recordHeaders, payload, isCompleted);
    var record = recordResult.record || {};
    var recordId = String(recordResult.recordId || record['학습기록ID'] || '').trim();

    /* WM_FINAL_CURRENT_PROGRESS_CORE4_REPAIR_20260820_V1
     * 학습기록_DB 완료 후 현재진행_DB 갱신 직전에 오류가 나면 학습기록은 완료인데
     * 현재진행은 STEP5/83%에 머물 수 있습니다. 완료 재요청 시 핵심 4개 값이 아직
     * 확정되지 않은 경우에만 공식 전환을 다시 계산해 현재진행_DB를 복구합니다. */
    if (isCompleted && recordResult.targetRow > 1 && isCompleteStatus(record['완료상태']) && String(record['완료'] || '').trim()) {
      var retrySetId = String(record['Set_ID'] || payload.Set_ID || payload.setId || '').trim();
      if (!wmFastIsCurrentProgressCompletionCommitted_(progressBefore.obj || {}, retrySetId)) {
        var retryTransition = wmFastResolveCompletionTransition_(payload, progressBefore.obj || {}, retrySetId);
        var retryProgressResult = wmFastWriteCurrentProgress_(ss, payload, recordId, true, retryTransition, record);
        if (!retryProgressResult || retryProgressResult.success !== true) {
          wmClearRuntimeCachesForStudent_(studentId);
          return outputResult(e, {
            success:false, alreadyCompleted:true, recordCompleted:true, currentProgressRepaired:false,
            recordId:recordId, setId:retrySetId,
            message:'학습기록은 완료되었으나 현재진행_DB 복구에 실패했습니다.',
            currentProgressUpdateError:String(retryProgressResult && retryProgressResult.message || '')
          });
        }
        /* WM_FINAL_RETURN_CURRENT_PROGRESS_DIRECT_V1_20260821
         * 현재진행_DB 완료 확정 뒤 1.학생관리_DB 현재세트를 다시 저장하지 않습니다.
         * Map/Study의 공식 진행 기준은 현재진행_DB이며, 완료 응답에서 그 1행을 그대로 반환합니다. */
        wmClearRuntimeCachesForStudent_(studentId);
        return outputResult(e, {
          success:true, alreadyCompleted:true, currentProgressRepaired:true,
          recordId:recordId, studentId:studentId, setId:retrySetId, status:'완료',
          transitionType:retryTransition.transitionType, nextOfficialSetId:retryTransition.nextSetId,
          completedRound:retryTransition.completedRound, targetRounds:retryTransition.targetRounds,
          transitionLevel:retryTransition.level, conditionCopied:retryTransition.conditionCopied,
          currentProgress:retryProgressResult.currentProgress || {},
          currentProgressCache:retryProgressResult.currentProgressCache || retryProgressResult.currentProgress || {},
          currentProgressRow:retryProgressResult.currentProgressRow || retryProgressResult.currentProgress || {},
          learningRecord:Object.assign({}, record || {})
        });
      }

      var retryResponse = wmFastInferCompletedRetryResponse_(progressBefore.obj || {}, record, payload);
      retryResponse.currentProgress = Object.assign({}, progressBefore.obj || {});
      retryResponse.currentProgressCache = Object.assign({}, progressBefore.obj || {});
      retryResponse.currentProgressRow = Object.assign({}, progressBefore.obj || {});
      retryResponse.learningRecord = Object.assign({}, record || {});
      wmClearRuntimeCachesForStudent_(studentId);
      return outputResult(e, retryResponse);
    }

    if (!isCompleted) {
      var stepProgress = wmFastWriteCurrentProgress_(ss, payload, recordId, false, null, record);
      wmClearRuntimeCachesForStudent_(studentId);
      return outputResult(e, {
        success:true,
        mode:'step_checkpoint_fast',
        recordId:recordId,
        studentId:studentId,
        setId:String(record['Set_ID'] || payload.Set_ID || payload.setId || '').trim(),
        status:'',
        currentProgressStep:String(payload.currentProgressStep || '').trim(),
        currentProgress:stepProgress && stepProgress.currentProgress ? stepProgress.currentProgress : {},
        currentProgressCache:stepProgress && (stepProgress.currentProgressCache || stepProgress.currentProgress) ? (stepProgress.currentProgressCache || stepProgress.currentProgress) : {},
        currentProgressRow:stepProgress && (stepProgress.currentProgressRow || stepProgress.currentProgress) ? (stepProgress.currentProgressRow || stepProgress.currentProgress) : {},
        learningRecord:Object.assign({}, record || {}),
        currentProgressUpdateError:stepProgress && stepProgress.success ? '' : String(stepProgress && stepProgress.message || '')
      });
    }

    var transition = wmFastResolveCompletionTransition_(payload, progressBefore.obj || {}, String(record['Set_ID'] || payload.Set_ID || payload.setId || '').trim());
    record['완료'] = wmFormatOfficialCompleteRound_(transition.recordRound || 1);
    record['완료상태'] = '완료';
    var finalRow = recordHeaders.map(function(header){ return record[header] !== undefined ? record[header] : ''; });

    /* TEST 완료는 현재 접속의 동일 학습기록ID 1행에서 완료 확정합니다. 재접속 후 TEST가 첫 완료단계인 경우에만 새 누적행을 생성합니다. */
    if (recordResult.needsInsert === true) {
      recordSheet.insertRowBefore(2);
      recordSheet.getRange(2, 1, 1, recordHeaders.length).setValues([finalRow]);
      recordResult.targetRow = 2;
    } else if (recordResult.targetRow > 1) {
      recordSheet.getRange(recordResult.targetRow, 1, 1, recordHeaders.length).setValues([finalRow]);
    } else {
      throw new Error('TEST 완료 새 누적행 저장 위치를 확정할 수 없습니다.');
    }

    var progressResult = wmFastWriteCurrentProgress_(ss, payload, recordId, true, transition, record);
    if (!progressResult || progressResult.success !== true) {
      wmClearRuntimeCachesForStudent_(studentId);
      return outputResult(e, {
        success:false, partialSaved:true, recordCompleted:true, currentProgressCompleted:false,
        recordId:recordId, studentId:studentId, setId:String(record['Set_ID'] || '').trim(), status:'완료',
        message:'학습기록은 완료되었으나 현재진행_DB 완료 확정에 실패했습니다.',
        currentProgressUpdateError:String(progressResult && progressResult.message || '')
      });
    }
    /* 현재진행_DB가 이미 공식 다음세트까지 확정했으므로 완료 hot path에서 추가 DB 저장을 하지 않습니다. */
    wmClearRuntimeCachesForStudent_(studentId);

    return outputResult(e, {
      success:true,
      mode:'set_complete_fast',
      recordId:recordId,
      studentId:studentId,
      setId:String(record['Set_ID'] || '').trim(),
      status:'완료',
      transitionType:transition.transitionType,
      nextOfficialSetId:transition.nextSetId,
      completedRound:transition.completedRound,
      targetRounds:transition.targetRounds,
      transitionLevel:transition.level,
      conditionCopied:transition.conditionCopied,
      currentProgress:progressResult.currentProgress || {},
      currentProgressCache:progressResult.currentProgressCache || progressResult.currentProgress || {},
      currentProgressRow:progressResult.currentProgressRow || progressResult.currentProgress || {},
      learningRecord:Object.assign({}, record || {}),
      currentProgressUpdateError:progressResult && progressResult.success ? '' : String(progressResult && progressResult.message || '')
    });
  } catch (err) {
    return outputResult(e, {success:false, message:'학습 완료 처리 오류', error:String(err && err.message ? err.message : err)});
  }
}

/* =========================================================
 * Word Mate LMS 학습기록 저장 API
 * action=saveLearningRecord
 * Study.html 학습 완료 후 2.학습기록_DB에 새 행을 추가합니다.
 * ========================================================= */
function saveLearningRecord(e) {
  try {
    var payload = readLearningPayload(e);

    if (!payload) {
      return outputResult(e, {
        success: false,
        message: '저장할 학습 payload가 없습니다.'
      });
    }

    var payloadStudentIdForSession = String(payload.studentId || payload.학생ID || '').trim().toUpperCase();
    var payloadSessionToken = String(payload.sessionToken || payload.현재세션 || payload.wmSessionToken || '').trim();
    if (payloadStudentIdForSession && payloadSessionToken && !wmIsStudentSessionValid_(payloadStudentIdForSession, payloadSessionToken)) {
      return outputResult(e, wmBuildSessionExpiredResponse_());
    }

    var ss = getLmsSpreadsheet_();
    if (!ss) {
      return outputResult(e, {
        success: false,
        message: '스프레드시트 연결 실패'
      });
    }

    var sheet = ss.getSheetByName('2.학습기록_DB');
    if (!sheet) {
      return outputResult(e, {
        success: false,
        message: '2.학습기록_DB 시트를 찾을 수 없습니다.'
      });
    }

    var values = sheet.getDataRange().getDisplayValues();
    var headers = values && values.length ? values[0].map(function(h){ return String(h || '').trim(); }) : [];

    if (!headers.length) {
      return outputResult(e, {
        success: false,
        message: '2.학습기록_DB 헤더가 없습니다.'
      });
    }

    headers = normalizeLearningRecordHeaders_(headers);

    var record = buildLearningRecordForSheet(payload, sheet, headers);
    var row = headers.map(function(header) {
      return record[header] !== undefined ? record[header] : '';
    });

    var targetRow = findLearningRecordRow_(sheet, headers, record['학습기록ID'], record['학생ID'], record['Set_ID']);
    record = wmApplyOfficialCompleteRoundToRecord_(sheet, headers, record, targetRow);
    row = headers.map(function(header) {
      return record[header] !== undefined ? record[header] : '';
    });
    var savedMode = 'insert';

    if (targetRow > 1) {
      var statusHeaderIndex = headers.indexOf('완료상태');
      var currentStatus = statusHeaderIndex >= 0
        ? String(sheet.getRange(targetRow, statusHeaderIndex + 1).getDisplayValue() || '').trim()
        : '';

      if (currentStatus === '완료') {
        savedMode = record['완료상태'] === '완료' ? 'skip_completed_duplicate' : 'skip_completed_locked';
      } else if (record['완료상태'] === '완료') {
        /* 기본 fallback도 TEST 완료는 새 줄 누적 스냅샷으로 저장합니다. */
        record = wmMergeExistingLearningRecordTimes_(sheet, headers, targetRow, record);
        record['학습기록ID'] = generateMonthlyLearningRecordId_(sheet, headers, new Date());
        row = headers.map(function(header) {
          return record[header] !== undefined ? record[header] : '';
        });
        sheet.insertRowBefore(2);
        sheet.getRange(2, 1, 1, headers.length).setValues([row]);
        savedMode = 'insert_complete_snapshot';
      } else {
        record = wmMergeExistingLearningRecordTimes_(sheet, headers, targetRow, record);
        row = headers.map(function(header) {
          return record[header] !== undefined ? record[header] : '';
        });
        sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
        savedMode = 'update';
      }
    } else {
      record = wmMergeLatestIncompleteSetTimesForNewRecord_(sheet, headers, record);
      row = headers.map(function(header) {
        return record[header] !== undefined ? record[header] : '';
      });
      /* WM_LEARNING_RECORD_LATEST_TOP_RESTORE_20260713
       * 신규 학습기록은 헤더 바로 아래 2행에 삽입하여 최근 기록을 항상 위에 둡니다. */
      sheet.insertRowBefore(2);
      sheet.getRange(2, 1, 1, headers.length).setValues([row]);
    }

    var officialTransition = WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1.TRANSITION.evaluate(record);
    record['__WM_OFFICIAL_TRANSITION_TYPE'] = officialTransition.transitionType;
    record['__WM_OFFICIAL_NEXT_SET_ID'] = officialTransition.nextSetId;

    /* WM_LEARNING_RECORD_FLUSH_BEFORE_PROGRESS_20260704_V1
     * 2.학습기록_DB 저장 직후 8.현재진행_DB 캐시를 만들기 전에
     * 방금 저장한 행이 getDataRange()에 반영되도록 강제 반영합니다.
     */
    try {
      SpreadsheetApp.flush();
    } catch (flushErrBasic) {}

    /* WM_CURRENT_PROGRESS_UPSERT_CALL_SAVELEARNINGRECORD_260618_V1
     * 기본 saveLearningRecord 경로에서도 8.현재진행_DB를 갱신합니다.
     * 기존에는 V2 경로에만 wmUpsertCurrentProgress_가 연결되어 있어
     * 실제 저장 경로가 기본 saveLearningRecord일 경우 현재진행_DB가 하단 append 상태로 남았습니다.
     */
    var currentProgressUpdateError = '';
    try {
      wmUpsertCurrentProgress_(ss, headers, record);
    } catch (ignoreProgressErrBasic) {
      currentProgressUpdateError = String(ignoreProgressErrBasic && ignoreProgressErrBasic.message ? ignoreProgressErrBasic.message : ignoreProgressErrBasic);
    }

    /* WM_STUDY_SAVE_SPEED_NO_DUPLICATE_PROGRESS_SORT_V1
     * wmUpsertCurrentProgress_가 최신 학생행을 이미 2행에 삽입하므로 재정렬하지 않습니다. */

    /* 최신 기록은 이미 2행에 삽입되므로 전체 학습기록 정렬을 생략합니다. */

    try {
      wmSyncStudentCurrentSetAfterOfficialCompletion_(
        record['학생ID'],
        record['완료상태'],
        officialTransition.nextSetId
      );
    } catch (ignoreOfficialCurrentSetErr) {
    }

    wmClearRuntimeCachesForStudent_(record['학생ID']);

    return outputResult(e, {
      success: true,
      message: '학습기록 저장 성공',
      mode: savedMode,
      recordId: record['학습기록ID'],
      studentId: record['학생ID'],
      setId: record['Set_ID'],
      status: record['완료상태'],
      transitionType: officialTransition.transitionType,
      nextOfficialSetId: officialTransition.nextSetId,
      completedRound: officialTransition.completedRound,
      targetRounds: officialTransition.targetRounds,
      transitionLevel: officialTransition.level,
      conditionCopied: officialTransition.conditionCopied,
      debugTotalTime: record['총소요시간'],
      debugStep1: record['Step1_총시간'],
      debugStep2: record['Step2_총시간'],
      debugStep3: record['Step3_총시간'],
      debugStep4: record['Step4_총시간'],
      debugBeforeTest: record['Step5_총시간'] || record['시험전재학습'],
      debugTestTime: record['Test_총시간'],
      currentProgressUpdateError: currentProgressUpdateError
    });

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: '학습기록 저장 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}


/* =========================================================
 * WM_SAVE_LEARNING_RECORD_OFFICIAL_TOTAL_V2_20260608
 * 총소요시간 전용 우회 저장 API입니다.
 * Study.html은 action=saveLearningRecordV2로 이 함수만 호출합니다.
 * 목적: 기존 saveLearningRecord 경로에서 payload.elapsedSeconds 또는 마지막 Step 시간이
 * 총소요시간으로 들어가는 문제를 차단하고, 공식 Step 컬럼 합산값을 총소요시간에 직접 씁니다.
 * 공식: 총소요시간 = Step1 + Step2 + Step3 + Step4 + Step5_총시간 + Test_총시간
 * ========================================================= */
function wmSaveLearningRecordOfficialTotalV2_(e) {
  try {
    var payload = readLearningPayload(e);

    if (!payload) {
      return outputResult(e, {
        success: false,
        message: '저장할 학습 payload가 없습니다. V2'
      });
    }

    var payloadStudentIdForSession = String(payload.studentId || payload.학생ID || '').trim().toUpperCase();
    var payloadSessionToken = String(payload.sessionToken || payload.현재세션 || payload.wmSessionToken || '').trim();
    if (payloadStudentIdForSession && payloadSessionToken) {
      var sessionStateForSave = getStudentSessionState_(payloadStudentIdForSession);
      if (sessionStateForSave.success) {
        var saveSessionValid = sessionStateForSave.currentSession === payloadSessionToken && String(sessionStateForSave.sessionStatus || '').toUpperCase() === 'LOGIN';
        if (!saveSessionValid) {
          return outputResult(e, wmBuildSessionExpiredResponse_());
        }
      }
    }

    var ss = getLmsSpreadsheet_();
    if (!ss) {
      return outputResult(e, {
        success: false,
        message: '스프레드시트 연결 실패 V2'
      });
    }

    var sheet = ss.getSheetByName('2.학습기록_DB');
    if (!sheet) {
      return outputResult(e, {
        success: false,
        message: '2.학습기록_DB 시트를 찾을 수 없습니다. V2'
      });
    }

    var values = sheet.getDataRange().getDisplayValues();
    var headers = values && values.length ? values[0].map(function(h){ return String(h || '').trim(); }) : [];

    if (!headers.length) {
      return outputResult(e, {
        success: false,
        message: '2.학습기록_DB 헤더가 없습니다. V2'
      });
    }

    headers = normalizeLearningRecordHeaders_(headers);

    var record = wmBuildLearningRecordOfficialTotalV2_(payload, sheet, headers);
    var row = headers.map(function(header) {
      return record[header] !== undefined ? record[header] : '';
    });

    var targetRow = findLearningRecordRow_(sheet, headers, record['학습기록ID'], record['학생ID'], record['Set_ID']);
    record = wmApplyOfficialCompleteRoundToRecord_(sheet, headers, record, targetRow, values);
    row = headers.map(function(header) {
      return record[header] !== undefined ? record[header] : '';
    });
    var savedMode = 'insert_v2';

    if (targetRow > 1) {
      var statusHeaderIndex = headers.indexOf('완료상태');
      var currentStatus = statusHeaderIndex >= 0
        ? String(sheet.getRange(targetRow, statusHeaderIndex + 1).getDisplayValue() || '').trim()
        : '';

      if (currentStatus === '완료') {
        savedMode = record['완료상태'] === '완료' ? 'skip_completed_duplicate_v2' : 'skip_completed_locked_v2';
      } else if (record['완료상태'] === '완료') {
        /* fallback 경로도 TEST 완료는 직전 누적행을 덮어쓰지 않고 새 완료 스냅샷으로 저장합니다. */
        record = wmMergeExistingLearningRecordTimes_(sheet, headers, targetRow, record);
        record['학습기록ID'] = generateMonthlyLearningRecordId_(sheet, headers, new Date());
        row = headers.map(function(header) {
          return record[header] !== undefined ? record[header] : '';
        });
        sheet.insertRowBefore(2);
        sheet.getRange(2, 1, 1, headers.length).setValues([row]);
        values.splice(1, 0, row.slice());
        savedMode = 'insert_complete_snapshot_v2';
      } else {
        record = wmMergeExistingLearningRecordTimes_(sheet, headers, targetRow, record);
        row = headers.map(function(header) {
          return record[header] !== undefined ? record[header] : '';
        });
        sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
        values[targetRow - 1] = row.slice();
        savedMode = 'update_v2';
      }
    } else {
      record = wmMergeLatestIncompleteSetTimesForNewRecord_(sheet, headers, record);
      row = headers.map(function(header) {
        return record[header] !== undefined ? record[header] : '';
      });
      /* WM_LEARNING_RECORD_LATEST_TOP_RESTORE_20260713
       * 신규 학습기록은 헤더 바로 아래 2행에 삽입하여 최근 기록을 항상 위에 둡니다. */
      sheet.insertRowBefore(2);
      sheet.getRange(2, 1, 1, headers.length).setValues([row]);
      values.splice(1, 0, row.slice());
    }

    var officialTransitionV2 = WM_CODEGS_LEVEL_TRANSITION_NOTICE_LOCK_V1.TRANSITION.evaluate(record);
    record['__WM_OFFICIAL_TRANSITION_TYPE'] = officialTransitionV2.transitionType;
    record['__WM_OFFICIAL_NEXT_SET_ID'] = officialTransitionV2.nextSetId;

    /* WM_FINAL_SAVE_REUSE_RECORD_VALUES_SPEED_V1_20260711
     * 최종 저장에서 이미 읽은 학습기록_DB 값에 방금 저장한 행을 메모리로 반영했습니다.
     * 현재진행_DB 계산 전 강제 flush와 동일 시트 전체 재조회를 생략합니다. */

    var currentProgressUpdateErrorV2 = '';
    var currentProgressUpdateResultV2 = null;
    try {
  currentProgressUpdateResultV2 = WM_STUDY_SPEED_SERVER_LOCK_V1.currentProgressFast(ss, headers, record, values);
} catch (ignoreProgressErr) {
  currentProgressUpdateErrorV2 = String(ignoreProgressErr && ignoreProgressErr.message ? ignoreProgressErr.message : ignoreProgressErr);
}

/* WM_STUDY_SAVE_SPEED_NO_DUPLICATE_PROGRESS_SORT_V1
 * wmUpsertCurrentProgress_가 최신 학생행을 이미 2행에 삽입하므로 재정렬하지 않습니다. */

/* 최신 기록은 이미 2행에 삽입되므로 전체 학습기록 정렬을 생략합니다. */

try {
  wmSyncStudentCurrentSetAfterOfficialCompletion_(
    record['학생ID'],
    record['완료상태'],
    officialTransitionV2.nextSetId || (currentProgressUpdateResultV2 && currentProgressUpdateResultV2.nextSetId)
  );
} catch (ignoreOfficialCurrentSetErrV2) {
}

    wmClearRuntimeCachesForStudent_(record['학생ID']);

    return outputResult(e, {
      success: true,
      message: '학습기록 저장 성공 V2',
      mode: savedMode,
      recordId: record['학습기록ID'],
      studentId: record['학생ID'],
      setId: record['Set_ID'],
      status: record['완료상태'],
      transitionType: officialTransitionV2.transitionType,
      nextOfficialSetId: officialTransitionV2.nextSetId,
      completedRound: officialTransitionV2.completedRound,
      targetRounds: officialTransitionV2.targetRounds,
      transitionLevel: officialTransitionV2.level,
      conditionCopied: officialTransitionV2.conditionCopied,
      debugSaveRoute: 'saveLearningRecordV2',
      debugTotalTime: record['총소요시간'],
      debugStep1: record['Step1_총시간'],
      debugStep2: record['Step2_총시간'],
      debugStep3: record['Step3_총시간'],
      debugStep4: record['Step4_총시간'],
      debugBeforeTest: record['Step5_총시간'] || record['시험전재학습'],
      debugTestTime: record['Test_총시간'],
      debugFormula: 'Step1+Step2+Step3+Step4+Step5_총시간+Test_총시간',
      currentProgressUpdateError: currentProgressUpdateErrorV2
    });

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: '학습기록 저장 서버 오류 V2',
      error: String(err && err.message ? err.message : err)
    });
  }
}

function wmResolveLearningRecordSetIdForSave_(payload, studentId) {
  /* WM_SAVE_RECORD_SET_ID_ROOT_FIX_20260624_V1
   * 학습기록_DB Set_ID는 반드시 실제 학습 중인 세트가 저장되어야 합니다.
   * 일부 Study payload가 오래된 setId(WM4-1-1 등)를 들고 들어오는 경우가 있어,
   * 1) payload의 여러 Set_ID 후보를 먼저 수집하고,
   * 2) 후보가 비어 있거나 현재진행_DB의 기록Set_ID보다 과거 레벨이면 기록Set_ID로 보정합니다.
   * 학생ID가 기준키이므로 학생별 최신 현재진행_DB만 사용합니다.
   */
  payload = payload || {};
  studentId = String(studentId || payload.studentId || payload.학생ID || '').trim().toUpperCase();

  function normalizeWmSet_(value) {
    var text = String(value || '').trim().toUpperCase();
    if (!text) return '';
    if (/^WM\d+-\d+-\d+$/.test(text)) return text;
    if (/^\d+-\d+-\d+$/.test(text)) return 'WM' + text;
    return '';
  }

  function serial_(setId) {
    var p = parseWordMateSetId(setId);
    if (!p) return 0;
    return Number(p.level || 0) * 10000 + Number(p.part || 0) * 100 + Number(p.set || 0);
  }

  var candidates = [
    payload.setId,
    payload.set_id,
    payload.Set_ID,
    payload.currentSetId,
    payload.currentSet,
    payload.wmCurrentSet,
    payload.initialSetId,
    payload.WM_INITIAL_SET_ID,
    payload['현재세트'],
    payload['기록Set_ID']
  ];

  var payloadSetId = '';
  for (var i = 0; i < candidates.length; i++) {
    payloadSetId = normalizeWmSet_(candidates[i]);
    if (payloadSetId) break;
  }

  var progressSetId = '';
  try {
    /* WM_SAVE_SET_RESOLVE_SINGLE_PROGRESS_20260704_V1
     * 저장 시 전체 현재진행맵 캐시를 쓰면 이전 세트가 남을 수 있어
     * 학생 1명의 최신 현재진행_DB 행만 직접 확인합니다.
     */
    progressSetId = normalizeWmSet_(wmGetCurrentProgressSetForStudent_(studentId));
  } catch (err) {
    progressSetId = '';
  }

  if (!payloadSetId && progressSetId) {
    return progressSetId;
  }

  if (payloadSetId && progressSetId) {
    var payloadParsed = parseWordMateSetId(payloadSetId);
    var progressParsed = parseWordMateSetId(progressSetId);
    if (payloadParsed && progressParsed) {
      var payloadSerial = serial_(payloadSetId);
      var progressSerial = serial_(progressSetId);
      if (Number(progressParsed.level || 0) > Number(payloadParsed.level || 0)) {
        return progressSetId;
      }
      if (progressSerial > payloadSerial && String(payloadSetId) === 'WM4-1-1') {
        return progressSetId;
      }
    }
  }

  return payloadSetId || 'WM5-1-1';
}

function wmNormalizeWrongWordListForRecord_(items) {
  var seen = {};
  var normalized = [];
  items = Array.isArray(items) ? items : [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i] || {};
    var word = String(item.word || item.영어 || '').trim();
    var key = word.toLowerCase();
    if (!key || seen[key]) continue;
    seen[key] = true;
    normalized.push({
      word: word,
      pos: String(item.pos || item.품사 || '').trim(),
      meaning: String(item.meaning || item.한글 || item.korean || '').trim()
    });
  }
  return normalized;
}

function wmBuildWrongWordsCellValue_(payload) {
  var source = payload && payload.wrongWords;
  if (!source || Array.isArray(source) || typeof source !== 'object') return '';
  var normalized = {
    version: 'WM_WRONG_WORDS_V1',
    engKoWrite: wmNormalizeWrongWordListForRecord_(source.engKoWrite),
    koEngWrite: wmNormalizeWrongWordListForRecord_(source.koEngWrite)
  };
  return JSON.stringify(normalized);
}

function wmBuildLearningRecordOfficialTotalV2_(payload, sheet, headers) {
  var now = new Date();
  var studentIdForProfile = String(payload.studentId || payload.학생ID || '').trim();
  var setId = wmResolveLearningRecordSetIdForSave_(payload, studentIdForProfile);
  var parsed = parseWordMateSetId(setId);

  var scores = Array.isArray(payload.scores) ? payload.scores : [];
  var scoreMap = extractScoreMap(scores);
  var finalDisplayScore = getFinalKoEngWriteScoreFromPayload_(payload, scoreMap);

  var completedSteps = Array.isArray(payload.completedSteps) ? payload.completedSteps : [];
  var officialTimeFields = wmBuildOfficialLearningTimeFieldsV2_(completedSteps, payload.officialTimeSummary);

  var recordId = String(payload.recordSessionId || payload.recordId || payload['학습기록ID'] || '').trim();
  var studentProfile = getStudentBasicInfoForMap_(studentIdForProfile);

  if (!isOfficialLearningRecordId_(recordId)) {
    recordId = generateMonthlyLearningRecordId_(sheet, headers, now);
  }

  return {
    '학습기록ID': recordId,
    '학습날짜': Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    '학생ID': studentIdForProfile,
    '학생이름': resolveStudentNameForLearningRecord_(studentProfile, payload, studentIdForProfile),
    '학교': String(payload.school || payload.학교 || studentProfile.school || '').trim(),
    '학년': String(payload.grade || payload.학년 || studentProfile.grade || '').trim(),
    'Class': String(studentProfile.className || payload.className || payload.Class || payload['반명'] || payload['반'] || '').trim(),
    '교사명': String(payload.teacherName || payload.교사명 || studentProfile.teacherName || '').trim(),
    'Set_ID': setId,
    '점수': finalDisplayScore,
    '총소요시간': officialTimeFields.totalTime,
    'Step1_총시간': officialTimeFields.step1,
    'Step2_총시간': officialTimeFields.step2,
    'Step3_총시간': officialTimeFields.step3,
    'Step4_총시간': officialTimeFields.step4,
    'Step5_총시간': officialTimeFields.beforeTest,
    '시험전재학습': officialTimeFields.beforeTest,
    'Test_총시간': officialTimeFields.testTime,
    '영한객관식': scoreMap.engKoChoice,
    '한영객관식': scoreMap.koEngChoice,
    '영한주관식': scoreMap.engKoWrite,
    '한영주관식': scoreMap.koEngWrite,
    '틀린단어': wmBuildWrongWordsCellValue_(payload),
    '완료': '',
    '완료상태': normalizeLearningStatus_(payload.status),
    '기기정보': normalizeDeviceInfo_(payload.userAgent || payload.deviceInfo || ''),
    '__WM_PAYLOAD_LAST_COMPLETED_STEP': payload.lastCompletedStep || '',
    '__WM_PAYLOAD_CURRENT_PROGRESS_STEP': payload.currentProgressStep || '',
    '__WM_PAYLOAD_STATUS': payload.status || '',
    '비고': parsed ? '' : 'Set_ID 파싱 확인 필요'
  };
}

function wmBuildOfficialLearningTimeFieldsV2_(completedSteps, summary) {
  var stepSeconds = {
    step1_flash_card: 0,
    step2_choose_meaning: 0,
    step3_choose_english: 0,
    step4_memory: 0,
    step6_last_flash_review: 0,
    test_all: 0
  };

  var stepKeyMap = {
    step1: 'step1_flash_card',
    step2: 'step2_choose_meaning',
    step3: 'step3_choose_english',
    step4: 'step4_memory',
    beforeTest: 'step6_last_flash_review',
    step5: 'step6_last_flash_review',
    testTime: 'test_all'
  };

  summary = summary || {};
  var summaryStepSeconds = summary.stepSeconds || {};

  for (var directKey in stepKeyMap) {
    if (Object.prototype.hasOwnProperty.call(stepKeyMap, directKey)) {
      var stepIdFromSummary = stepKeyMap[directKey];
      var directValue = Number(summary[directKey] || summaryStepSeconds[stepIdFromSummary] || 0);
      if (isFinite(directValue) && directValue > 0) {
        stepSeconds[stepIdFromSummary] = Math.max(0, Math.round(directValue));
      }
    }
  }

  completedSteps = Array.isArray(completedSteps) ? completedSteps : [];
  for (var i = 0; i < completedSteps.length; i++) {
    var step = completedSteps[i] || {};
    var stepId = String(step.stepId || '').trim();
    if (!Object.prototype.hasOwnProperty.call(stepSeconds, stepId)) {
      continue;
    }

    var seconds = Number(step.elapsedSeconds || 0);
    if (!isFinite(seconds) || seconds <= 0) {
      seconds = parseTimeTextToSeconds_(step.elapsedText || step.timeText || '');
    }

    if (isFinite(seconds) && seconds > 0) {
      /* 같은 Step이 중복 저장되면 마지막 유효값 하나만 공식값으로 사용합니다. */
      stepSeconds[stepId] = Math.max(0, Math.round(seconds));
    }
  }

  var totalSeconds = stepSeconds.step1_flash_card
    + stepSeconds.step2_choose_meaning
    + stepSeconds.step3_choose_english
    + stepSeconds.step4_memory
    + stepSeconds.step6_last_flash_review
    + stepSeconds.test_all;

  return {
    step1: formatSecondsForSheet(stepSeconds.step1_flash_card),
    step2: formatSecondsForSheet(stepSeconds.step2_choose_meaning),
    step3: formatSecondsForSheet(stepSeconds.step3_choose_english),
    step4: formatSecondsForSheet(stepSeconds.step4_memory),
    step5: formatSecondsForSheet(stepSeconds.step6_last_flash_review),
    beforeTest: formatSecondsForSheet(stepSeconds.step6_last_flash_review),
    testTime: formatSecondsForSheet(stepSeconds.test_all),
    totalTime: formatSecondsForSheet(totalSeconds),
    totalSeconds: totalSeconds
  };
}

function sortLearningRecordSheetByDateDesc_(sheet, headers) {
  /* WM_LEARNING_RECORD_DATE_DESC_SORT_V1
   * 2.학습기록_DB는 최근 학습기록이 상단, 이전 기록이 하단에 오도록
   * 저장 직후 학습날짜 기준 내림차순으로 정렬합니다.
   */
  try {
    var idxDate = headers.indexOf('학습날짜');
    var lastRow = sheet.getLastRow();
    if (idxDate < 0 || lastRow < 3) {
      return;
    }

    sheet.getRange(2, 1, lastRow - 1, headers.length).sort({
      column: idxDate + 1,
      ascending: false
    });
  } catch (err) {}
}

function readLearningPayload(e) {
  var text = '';

  if (e && e.parameter && e.parameter.payload) {
    text = String(e.parameter.payload || '');
  }

  if (!text && e && e.postData && e.postData.contents) {
    text = String(e.postData.contents || '');
  }

  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function buildLearningRecordForSheet(payload, sheet, headers) {
  var now = new Date();
  var studentIdForProfile = String(payload.studentId || payload.학생ID || '').trim();
  var setId = wmResolveLearningRecordSetIdForSave_(payload, studentIdForProfile);
  var parsed = parseWordMateSetId(setId);

  var scores = Array.isArray(payload.scores) ? payload.scores : [];
  var scoreMap = extractScoreMap(scores);

  /* WM_FINAL_SCORE_LAST_KO_ENG_WRITE_ONLY_V2
   * 최종 점수는 평균점수가 아니라 마지막 시험인 한영주관식 점수만 기록합니다.
   * title 매칭이 흔들려도 scores 배열 끝에서 주관식+한영 항목을 한 번 더 찾습니다.
   */
  var finalDisplayScore = getFinalKoEngWriteScoreFromPayload_(payload, scoreMap);
  var completedSteps = Array.isArray(payload.completedSteps) ? payload.completedSteps : [];
  var normalizedCompletedSteps = normalizeCompletedStepElapsedSeconds_(completedSteps);

  var recordId = String(payload.recordSessionId || payload.recordId || payload['학습기록ID'] || '').trim();

  var studentProfile = getStudentBasicInfoForMap_(studentIdForProfile);

  /* WM_RECORD_ID_R_YY_MM_SEQ_V1
   * 확정 공식: 학습기록ID = R + YY + MM + 3자리 순번
   * 예: R2606001 = R + 2026년(26) + 6월(06) + 해당 월 001번
   * 해당 월 기존 최대 순번을 찾아 +1 합니다. 한 달 최대 999개입니다.
   */
  if (!isOfficialLearningRecordId_(recordId)) {
    recordId = generateMonthlyLearningRecordId_(sheet, headers, now);
  }

  /* WM_TOTAL_TIME_AND_TEST_TIME_RULE_20260608_V2
   * 총소요시간은 DB에 실제로 기록되는 공식 Step 컬럼값의 합산입니다.
   * Step1 + Step2 + Step3 + Step4 + Step5_총시간 + Test_총시간만 더합니다.
   * 중복/재진입 payload가 와도 각 공식 Step은 마지막 유효값 1개만 사용합니다.
   */
  var officialTimeFields = buildOfficialLearningTimeFieldsFromPayloadSummary_(payload.officialTimeSummary)
    || buildOfficialLearningTimeFields_(normalizedCompletedSteps);

  return {
    '학습기록ID': recordId,
    '학습날짜': Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    '학생ID': studentIdForProfile,
    '학생이름': resolveStudentNameForLearningRecord_(studentProfile, payload, studentIdForProfile),
    '학교': String(payload.school || payload.학교 || studentProfile.school || '').trim(),
    '학년': String(payload.grade || payload.학년 || studentProfile.grade || '').trim(),
    'Class': String(studentProfile.className || payload.className || payload.Class || payload['반명'] || payload['반'] || '').trim(),
    '교사명': String(payload.teacherName || payload.교사명 || studentProfile.teacherName || '').trim(),
    'Set_ID': setId,
    '점수': finalDisplayScore,
    '총소요시간': officialTimeFields.totalTime,
    'Step1_총시간': officialTimeFields.step1,
    'Step2_총시간': officialTimeFields.step2,
    'Step3_총시간': officialTimeFields.step3,
    'Step4_총시간': officialTimeFields.step4,
    'Step5_총시간': officialTimeFields.beforeTest,
    '시험전재학습': officialTimeFields.beforeTest,
    'Test_총시간': officialTimeFields.testTime,
    '영한객관식': scoreMap.engKoChoice,
    '한영객관식': scoreMap.koEngChoice,
    '영한주관식': scoreMap.engKoWrite,
    '한영주관식': scoreMap.koEngWrite,
    '완료': '',
    '완료상태': normalizeLearningStatus_(payload.status),
    '기기정보': normalizeDeviceInfo_(payload.userAgent || payload.deviceInfo || ''),
    '__WM_PAYLOAD_LAST_COMPLETED_STEP': payload.lastCompletedStep || '',
    '__WM_PAYLOAD_CURRENT_PROGRESS_STEP': payload.currentProgressStep || '',
    '__WM_PAYLOAD_STATUS': payload.status || '',
    '비고': parsed ? '' : 'Set_ID 파싱 확인 필요'
  };
}


function wmNormalizeOfficialScoreValue_(value) {
  /* WM_SCORE_VALUE_GUARD_20260707_V1
   * 최근점수/점수에는 레벨명(예: 5레벨)이 들어가면 안 됩니다.
   * 숫자 점수 또는 퍼센트 점수만 인정하고, 그 외 문자열은 빈값으로 처리합니다.
   */
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return '';
  if (/레벨|LEVEL|WM\d+-\d+-\d+/i.test(text)) return '';
  var match = text.match(/^(\d{1,3})(?:\.\d+)?\s*(?:점|%)?$/);
  if (!match) return '';
  var n = Number(match[1]);
  if (!isFinite(n) || n < 0 || n > 100) return '';
  return String(Math.round(n));
}

function getFinalKoEngWriteScoreFromPayload_(payload, scoreMap) {
  /* WM_FINAL_SCORE_LAST_TEST_SCORE_V3_HELPER
   * 최종 점수는 학습모드에 따라 마지막으로 실시한 본시험 점수를 저장합니다.
   * - 객관식(2개): 한영객관식
   * - 객관식+영한: 영한주관식
   * - 객관식+한영: 한영주관식
   * - 기본전체: 한영주관식
   */
  var directFinal = wmNormalizeOfficialScoreValue_(payload && (payload.finalScore || payload.officialScore || payload.score));
  if (directFinal) {
    return directFinal;
  }

  var finalTestScore = payload && payload.finalTestScore ? payload.finalTestScore : null;
  var finalPercent = wmNormalizeOfficialScoreValue_(finalTestScore && finalTestScore.percent);
  if (finalPercent) {
    return finalPercent;
  }

  var scores = payload && Array.isArray(payload.scores) ? payload.scores : [];
  for (var i = scores.length - 1; i >= 0; i--) {
    var item = scores[i] || {};
    var percent = wmNormalizeOfficialScoreValue_(item.percent);
    if (percent) {
      return percent;
    }
  }

  return '';
}

function resolveStudentNameForLearningRecord_(studentProfile, payload, studentId) {
  /* WM_RECORD_STUDENT_NAME_DB_FIRST_V1
   * 학습기록_DB의 학생이름은 학생ID나 임시값이 아니라 실제 학생이름을 기록합니다.
   * 1.학생관리_DB 기준 이름을 최우선으로 사용하고, 없을 때만 payload 값을 사용합니다.
   */
  var profileName = String(studentProfile && studentProfile.studentName || '').trim();
  if (profileName) {
    return profileName;
  }

  var rawName = String((payload && (payload.studentName || payload['학생이름'] || payload.name)) || '').trim();
  var sid = String(studentId || (payload && payload.studentId) || (payload && payload['학생ID']) || '').trim().toUpperCase();

  if (!rawName || rawName === '학생' || rawName.toUpperCase() === sid) {
    return '';
  }

  return rawName;
}

function sumCompletedStepElapsedSeconds_(completedSteps) {
  /* WM_TOTAL_TIME_STEP_UNIT_SUM_V5
   * Study.html에서 Step별 elapsedSeconds가 누적시간으로 들어와도
   * 각 Step 시간을 개별시간으로 환산한 뒤 모두 더해 총소요시간을 기록합니다.
   */
  return sumNormalizedCompletedStepElapsedSeconds_(normalizeCompletedStepElapsedSeconds_(completedSteps));
}

function sumNormalizedCompletedStepElapsedSeconds_(normalizedSteps) {
  var total = 0;
  normalizedSteps = Array.isArray(normalizedSteps) ? normalizedSteps : [];

  for (var i = 0; i < normalizedSteps.length; i++) {
    var seconds = Number(normalizedSteps[i].__unitElapsedSeconds || 0);
    if (isFinite(seconds) && seconds > 0) {
      total += seconds;
    }
  }

  return Math.max(0, Math.round(total));
}

function buildOfficialLearningTimeFields_(normalizedSteps) {
  /* WM_TOTAL_TIME_OFFICIAL_COLUMNS_SUM_V1
   * 확정 공식 강제 적용:
   * 총소요시간 = Step1_총시간 + Step2_총시간 + Step3_총시간 + Step4_총시간 + Step5_총시간 + Test_총시간.
   * payload에 중복 Step이 있으면 마지막 유효 Step 시간 1개만 사용합니다.
   * 이 함수의 결과값을 DB 컬럼과 총소요시간에 동시에 사용하므로 표시값과 합산값이 어긋나지 않습니다.
   */
  normalizedSteps = Array.isArray(normalizedSteps) ? normalizedSteps : [];

  var officialIds = [
    'step1_flash_card',
    'step2_choose_meaning',
    'step3_choose_english',
    'step4_memory',
    'step6_last_flash_review',
    'test_all'
  ];

  var secondsByStep = {};

  for (var i = 0; i < normalizedSteps.length; i++) {
    var step = normalizedSteps[i] || {};
    var stepId = String(step.stepId || '').trim();
    if (officialIds.indexOf(stepId) === -1) {
      continue;
    }

    var seconds = Number(step.__unitElapsedSeconds || 0);
    if (!isFinite(seconds) || seconds <= 0) {
      seconds = parseTimeTextToSeconds_(step.elapsedText || step.timeText || '');
    }

    if (isFinite(seconds) && seconds > 0) {
      secondsByStep[stepId] = Math.max(0, Math.round(seconds));
    }
  }

  function getSeconds(stepId) {
    var value = Number(secondsByStep[stepId] || 0);
    return isFinite(value) && value > 0 ? Math.round(value) : 0;
  }

  var step1 = getSeconds('step1_flash_card');
  var step2 = getSeconds('step2_choose_meaning');
  var step3 = getSeconds('step3_choose_english');
  var step4 = getSeconds('step4_memory');
  var beforeTest = getSeconds('step6_last_flash_review');
  var testTime = getSeconds('test_all');
  var total = step1 + step2 + step3 + step4 + beforeTest + testTime;

  return {
    step1: formatSecondsForSheet(step1),
    step2: formatSecondsForSheet(step2),
    step3: formatSecondsForSheet(step3),
    step4: formatSecondsForSheet(step4),
    beforeTest: formatSecondsForSheet(beforeTest),
    testTime: formatSecondsForSheet(testTime),
    totalTime: formatSecondsForSheet(total),
    totalSeconds: total
  };
}

function buildOfficialLearningTimeFieldsFromPayloadSummary_(summary) {
  /* WM_CODEGS_USE_STUDY_OFFICIAL_TIME_SUMMARY_20260608_V1
   * Study.html이 payload에 직접 실어 보낸 공식 Step 시간 합산값을 최우선 사용합니다.
   * 이 값이 있으면 payload.elapsedSeconds/마지막 Step 시간은 총소요시간에 사용하지 않습니다. */
  summary = summary || {};
  var stepSeconds = summary.stepSeconds || {};

  function pick(key, altKey) {
    var value = Number(summary[key] || stepSeconds[altKey] || 0);
    return isFinite(value) && value > 0 ? Math.round(value) : 0;
  }

  var step1 = pick('step1', 'step1_flash_card');
  var step2 = pick('step2', 'step2_choose_meaning');
  var step3 = pick('step3', 'step3_choose_english');
  var step4 = pick('step4', 'step4_memory');
  var beforeTest = pick('beforeTest', 'step6_last_flash_review');
  var testTime = pick('testTime', 'test_all');
  var total = Number(summary.totalSeconds || 0);

  if (!isFinite(total) || total <= 0) {
    total = step1 + step2 + step3 + step4 + beforeTest + testTime;
  }

  if (!isFinite(total) || total <= 0) {
    return null;
  }

  return {
    step1: formatSecondsForSheet(step1),
    step2: formatSecondsForSheet(step2),
    step3: formatSecondsForSheet(step3),
    step4: formatSecondsForSheet(step4),
    beforeTest: formatSecondsForSheet(beforeTest),
    testTime: formatSecondsForSheet(testTime),
    totalTime: formatSecondsForSheet(total),
    totalSeconds: Math.round(total)
  };
}

function normalizeCompletedStepElapsedSeconds_(completedSteps) {
  /* WM_STEP_TIME_DIRECT_UNIT_SUM_V6
   * 확정 공식: Study.html은 각 Step의 개별 elapsedSeconds를 전달합니다.
   * 따라서 Code.gs는 누적형으로 오판해서 빼기 계산하지 않고,
   * 전달된 Step별 개별시간을 그대로 사용합니다.
   * 총소요시간 = Step1 + Step2 + Step3 + Step4 + Step5_총시간 + Test_총시간.
   */
  if (!Array.isArray(completedSteps)) {
    return [];
  }

  var list = [];
  for (var i = 0; i < completedSteps.length; i++) {
    var source = completedSteps[i] || {};
    var copied = {};
    for (var key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        copied[key] = source[key];
      }
    }
    var raw = Number(source.elapsedSeconds || 0);
    var parsedTextSeconds = parseTimeTextToSeconds_(source.elapsedText || source.timeText || '');
    var unitSeconds = isFinite(raw) && raw > 0 ? Math.max(0, Math.round(raw)) : parsedTextSeconds;
    copied.__originalIndex = i;
    copied.__rawElapsedSeconds = isFinite(raw) ? Math.max(0, Math.round(raw)) : 0;
    copied.__parsedElapsedTextSeconds = parsedTextSeconds;
    copied.__unitElapsedSeconds = unitSeconds;
    list.push(copied);
  }

  return list;
}

function extractScoreMap(scores) {
  var result = {
    engKoChoice: '',
    koEngChoice: '',
    engKoWrite: '',
    koEngWrite: ''
  };

  for (var i = 0; i < scores.length; i++) {
    var item = scores[i] || {};
    var title = String(item.title || '').replace(/<br\s*\/?>/gi, ' ');
    var percent = item.percent !== undefined ? String(item.percent) : '';

    if (!percent) {
      continue;
    }

    if (title.indexOf('객관식 영한') !== -1) {
      result.engKoChoice = percent;
    } else if (title.indexOf('객관식 한영') !== -1) {
      result.koEngChoice = percent;
    } else if (title.indexOf('주관식 영한') !== -1) {
      result.engKoWrite = percent;
    } else if (title.indexOf('주관식 한영') !== -1) {
      result.koEngWrite = percent;
    }
  }

  return result;
}

function calculateAverageScore(scores) {
  if (!Array.isArray(scores) || !scores.length) {
    return '';
  }

  var total = 0;
  var count = 0;

  for (var i = 0; i < scores.length; i++) {
    var value = Number(scores[i] && scores[i].percent);
    if (isFinite(value)) {
      total += value;
      count += 1;
    }
  }

  if (!count) {
    return '';
  }

  return Math.round(total / count);
}

function parseTimeTextToSeconds_(value) {
  /* WM_TIME_TEXT_PARSE_FOR_TOTAL_V1
   * 진행 재진입 시 Code.gs가 기존 DB의 '1분 20초' 같은 표시값을 completedSteps.elapsedText로 되돌려줍니다.
   * 이 경우 elapsedSeconds가 0이므로, 총소요시간 합산을 위해 표시 시간을 초 단위로 복원합니다.
   */
  var text = String(value || '').trim();
  if (!text) {
    return 0;
  }

  var total = 0;
  var hourMatch = text.match(/(\d+)\s*시간/);
  var minMatch = text.match(/(\d+)\s*분/);
  var secMatch = text.match(/(\d+)\s*초/);

  if (hourMatch) total += Number(hourMatch[1]) * 3600;
  if (minMatch) total += Number(minMatch[1]) * 60;
  if (secMatch) total += Number(secMatch[1]);

  if (total > 0) {
    return total;
  }

  var colonMatch = text.match(/^(\d+):(\d+)(?::(\d+))?$/);
  if (colonMatch) {
    if (colonMatch[3] !== undefined) {
      return Number(colonMatch[1]) * 3600 + Number(colonMatch[2]) * 60 + Number(colonMatch[3]);
    }
    return Number(colonMatch[1]) * 60 + Number(colonMatch[2]);
  }

  var numeric = Number(text.replace(/[^0-9.]/g, ''));
  return isFinite(numeric) ? Math.max(0, Math.round(numeric)) : 0;
}

function formatSecondsForSheet(seconds) {
  var sec = Number(seconds || 0);
  if (!isFinite(sec) || sec <= 0) {
    return '';
  }

  sec = Math.round(sec);
  var hour = Math.floor(sec / 3600);
  var remain = sec % 3600;
  var min = Math.floor(remain / 60);
  var rest = remain % 60;

  if (hour > 0) {
    return hour + '시간 ' + min + '분 ' + rest + '초';
  }

  return min + '분 ' + rest + '초';
}

function wmGetNextPlainSetIdForCurrentProgress_(setId) {
  var plain = normalizeLevelPlainSetId_(setId);
  var parts = plain.match(/^(\d+)-(\d+)-(\d+)$/);
  if (!parts) return plain;

  var level = Number(parts[1]);
  var part = Number(parts[2]);
  var set = Number(parts[3]);

  var maxSet = getLearningMapPartSetCount_(level, part);
  if (maxSet && set < maxSet) {
    return level + '-' + part + '-' + (set + 1);
  }

  var partCounts = getLearningMapLevelPartSetCounts_(level);
  if (part < partCounts.length) {
    return level + '-' + (part + 1) + '-1';
  }

  var nextLevel = level + 1;
  if (nextLevel <= 13) {
    return nextLevel + '-1-1';
  }

  return plain;
}

function wmNormalizeCurrentProgressCompletedStep_(step) {
  var text = String(step || '').trim().toUpperCase();
  if (text === 'TEST' || text === 'STEP6' || text === 'TEST_ALL') return 'TEST';
  if (text === 'BEFORE_TEST' || text === 'STEP6_LAST_FLASH_REVIEW' || text === 'LAST_FLASH_REVIEW') return 'STEP5';
  if (text === 'COMPLETE' || text === 'COMPLETED') return 'TEST';
  if (text === 'STEP1' || text === 'STEP1_FLASH_CARD') return 'STEP1';
  if (text === 'STEP2' || text === 'STEP2_CHOOSE_MEANING') return 'STEP2';
  if (text === 'STEP3' || text === 'STEP3_CHOOSE_ENGLISH') return 'STEP3';
  if (text === 'STEP4' || text === 'STEP4_MEMORY') return 'STEP4';
  if (text === 'STEP5') return 'STEP5';
  return '';
}

function wmGetCurrentStepAfterCompletedStep_(completedStep) {
  var text = wmNormalizeCurrentProgressCompletedStep_(completedStep);
  if (text === 'STEP1') return 'STEP2';
  if (text === 'STEP2') return 'STEP3';
  if (text === 'STEP3') return 'STEP4';
  if (text === 'STEP4') return 'STEP5';
  if (text === 'STEP5') return 'TEST';
  if (text === 'TEST') return 'STEP1';
  return 'STEP1';
}

function wmGetFutureStepAfterCurrentStep_(currentStep) {
  var text = String(currentStep || '').trim().toUpperCase();
  if (text === 'STEP1') return 'STEP2';
  if (text === 'STEP2') return 'STEP3';
  if (text === 'STEP3') return 'STEP4';
  if (text === 'STEP4') return 'STEP5';
  if (text === 'STEP5') return 'TEST';
  if (text === 'TEST' || text === 'STEP6') return 'STEP1';
  return 'STEP1';
}

function wmRuntimeStepFromCurrentProgress_(step) {
  var text = String(step || '').trim().toUpperCase();
  if (text === 'STEP6') return 'TEST';
  return text;
}

function wmFormatCurrentProgressSetIdLikeRecord_(plainSetId, recordSetId) {
  var plain = normalizeLevelPlainSetId_(plainSetId || recordSetId);
  if (!plain) return String(recordSetId || '').trim();
  return String(recordSetId || '').trim().toUpperCase().indexOf('WM') === 0 ? ('WM' + plain) : plain;
}

function wmBuildCurrentProgressStepStateFromLearningRecord_(record) {
  /* WM_CURRENT_PROGRESS_STEP_ENGINE_20260624_V4
   * 공식:
   * 1차 기준은 2.학습기록_DB의 record['Set_ID']입니다.
   * 8.현재진행_DB는 학습기록 Set_ID를 원본 Set_ID로 유지하고,
   * 완료Step(과거) → 현재Step(지금) → 기록Set_ID(현재Step 소속 세트)를 기록합니다.
   * 예: WM6-1-2 STEP5 완료 → Set_ID WM6-1-2 / 현재Step TEST / 기록Set_ID WM6-1-2.
   * 예: WM6-1-2 TEST 완료 → Set_ID WM6-1-2 / 현재Step STEP1 / 기록Set_ID WM6-1-3.
   */
  record = record || {};

  var learningRecordSetId = String(record['Set_ID'] || '').trim();
  var progress = deriveProgressFromLearningRecord_(record);
  var isCompleted = normalizeLearningStatus_(record['완료상태']) === '완료';

  /* WM_CURRENT_PROGRESS_PAYLOAD_STEP_HINT_20260704_V1
   * Study.html은 저장 payload에 lastCompletedStep/currentProgressStep을 같이 보냅니다.
   * Step 시간 컬럼이 빈값으로 들어오는 순간에도 현재진행_DB가 0%로 떨어지지 않도록
   * payload 기반 Step 힌트를 내부 필드로 전달받아 우선 보정합니다.
   */
  var payloadCompletedStep = wmNormalizeCurrentProgressCompletedStep_(
    record['__WM_PAYLOAD_LAST_COMPLETED_STEP'] || record['lastCompletedStep'] || record['currentProgressStep'] || ''
  );
  var completedStep = wmNormalizeCurrentProgressCompletedStep_(progress.lastCompletedStep || '') || payloadCompletedStep;
  var currentStep = 'STEP1';
  var recordStepSetId = learningRecordSetId;

  if (isCompleted) {
    completedStep = 'TEST';
    currentStep = 'STEP1';
    var officialNextSetId = String(record['__WM_OFFICIAL_NEXT_SET_ID'] || '').trim();
    var officialTransitionType = String(record['__WM_OFFICIAL_TRANSITION_TYPE'] || '').trim();
    if (officialNextSetId) {
      recordStepSetId = wmFormatCurrentProgressSetIdLikeRecord_(officialNextSetId, learningRecordSetId);
    } else if (officialTransitionType === 'COURSE_COMPLETE') {
      recordStepSetId = learningRecordSetId;
    } else {
      recordStepSetId = wmFormatCurrentProgressSetIdLikeRecord_(
        wmGetNextPlainSetIdForCurrentProgress_(learningRecordSetId),
        learningRecordSetId
      );
    }
  } else {
    currentStep = wmGetCurrentStepAfterCompletedStep_(completedStep);
    recordStepSetId = wmFormatCurrentProgressSetIdLikeRecord_(learningRecordSetId, learningRecordSetId);
  }

  return {
    Set_ID: learningRecordSetId,
    완료Step: completedStep,
    현재Step: currentStep,
    기록Set_ID: recordStepSetId,
    learningRecordSetId: learningRecordSetId
  };
}



function wmCalculateSetProgressPercentFromStepState_(stepState, record) {
  /* WM_SET_PROGRESS_PERCENT_20260624_V1
   * 세트진행률은 현재 학습 세트의 Step 진행 상태입니다.
   * 완료된 세트(TEST 완료)는 다음 기록Set_ID STEP1 상태가 되므로 0으로 표시합니다.
   */
  stepState = stepState || {};
  record = record || {};

  if (normalizeLearningStatus_(record['완료상태']) === '완료') {
    return 0;
  }

  var completedStep = wmNormalizeCurrentProgressCompletedStep_(stepState.완료Step || '');

  /* WM_SET_PROGRESS_TIME_FALLBACK_20260704_V1
   * 완료Step 계산이 비어도 실제 Step 시간 컬럼이 저장되어 있으면 그 값을 기준으로
   * 세트진행률을 복구합니다. Step4 저장 후 0%로 남는 현상을 막습니다.
   */
  if (!completedStep) {
    completedStep = wmNormalizeCurrentProgressCompletedStep_(record['__WM_PAYLOAD_LAST_COMPLETED_STEP'] || record['lastCompletedStep'] || record['currentProgressStep'] || '');
  }

  if (!completedStep) {
    if (String(record['Test_총시간'] || '').trim()) completedStep = 'TEST';
    else if (String(record['Step5_총시간'] || record['시험전재학습'] || record['시험전학습'] || '').trim()) completedStep = 'STEP5';
    else if (String(record['Step4_총시간'] || '').trim()) completedStep = 'STEP4';
    else if (String(record['Step3_총시간'] || '').trim()) completedStep = 'STEP3';
    else if (String(record['Step2_총시간'] || '').trim()) completedStep = 'STEP2';
    else if (String(record['Step1_총시간'] || '').trim()) completedStep = 'STEP1';
  }

  if (completedStep === 'TEST') return 100;
  if (completedStep === 'STEP5') return 83;
  if (completedStep === 'STEP4') return 67;
  if (completedStep === 'STEP3') return 50;
  if (completedStep === 'STEP2') return 33;
  if (completedStep === 'STEP1') return 17;
  return 0;
}

function wmSafeIsCompleteStatusForCurrentProgress_(status) {
  var text = String(status || '').trim();
  if (!text) return false;
  try {
    if (typeof isCompleteStatus === 'function' && isCompleteStatus(text)) return true;
  } catch (err1) {}
  try {
    if (typeof normalizeLearningStatus_ === 'function' && normalizeLearningStatus_(text) === '완료') return true;
  } catch (err2) {}
  return text === '완료' || text.toUpperCase() === 'COMPLETE';
}

function wmExtractScoreForCurrentProgress_(record) {
  record = record || {};
  return wmNormalizeOfficialScoreValue_(record['점수'])
    || wmNormalizeOfficialScoreValue_(record['한영주관식'])
    || wmNormalizeOfficialScoreValue_(record['최근점수'])
    || '';
}

function wmBuildRecent30JsonForCurrentProgress_(items) {
  var list = Array.isArray(items) ? items.slice(0, 30) : [];
  try {
    return JSON.stringify(list);
  } catch (err) {
    return '[]';
  }
}


/* WM_CURRENT_PROGRESS_LEVEL_HISTORY_CACHE_20260707_V3
 * 2.학습기록_DB 기록을 레벨별 순차완주회차로 집계해
 * 8.현재진행_DB 신규 컬럼(순차완주세트/순차완주회차/히스토리레벨/히스토리횟수)에 저장합니다.
 *
 * [순차완주세트 LOCK]
 * - 현재 회차에서 앞에서부터 순서대로 완료한 세트 수입니다.
 * - 현재 학습중 세트는 포함하지 않습니다.
 * - 중간 미완료 세트가 나오면 그 지점에서 카운트를 중단합니다.
 * - 완료세트수는 순서와 무관한 레벨 내 완료 세트 총수이고, 순차완주세트와 다릅니다.
 *
 * [색상 공식 LOCK]
 * - 색상은 레벨완료조건 + 순차완주회차 + 순차완주세트 공식에 맞을 때만 변합니다.
 * - 순차완주회차: 학습없음=0, 1회차 진행중=1, 2회차 진행중=2, 3회차 진행중=3.
 */
function wmBuildCurrentProgressLevelHistoryCache_(completedLogs, currentLevel, historyLogs) {
  var result = {
    레벨완료횟수: 0,
    순차완주세트: '',
    순차완주회차: 0,
    히스토리레벨: '',
    히스토리횟수: '',
    완료된레벨: '',
    완료횟수: ''
  };

  completedLogs = Array.isArray(completedLogs) ? completedLogs.slice() : [];
  historyLogs = Array.isArray(historyLogs) ? historyLogs.slice() : completedLogs.slice();
  currentLevel = Number(currentLevel || 0);

  completedLogs.sort(function(a, b) {
    var at = Number(a && a.sortTime || 0);
    var bt = Number(b && b.sortTime || 0);
    if (at !== bt) return at - bt;
    return Number(a && a.rowIndex || 0) - Number(b && b.rowIndex || 0);
  });

  var levelHasRecord = {};

  /* WM_LEVEL_HISTORY_ALL_STUDIED_LEVELS_20260706_V1
   * 학습히스토리 버튼 개수는 레벨완료 개수가 아니라
   * 2.학습기록_DB에 존재하는 모든 학습 레벨 + 현재진행_DB 현재레벨 기준입니다.
   * 회차/색상 계산은 기존 완료기록(completedLogs) 순차완주 공식으로만 판단합니다.
   */
  for (var i = 0; i < historyLogs.length; i++) {
    var plain = normalizeLevelPlainSetId_(
      (historyLogs[i] && (historyLogs[i].setId || historyLogs[i].Set_ID)) || ''
    );
    if (!plain) continue;
    var level = Number(String(plain).split('-')[0] || 0);
    if (level >= 3 && level <= 13) levelHasRecord[level] = true;
  }
  if (currentLevel >= 3 && currentLevel <= 13) levelHasRecord[currentLevel] = true;

  var levels = Object.keys(levelHasRecord).map(function(level){ return Number(level); }).filter(function(level){
    return level >= 3 && level <= 13;
  }).sort(function(a,b){ return a-b; });

  var roundsByLevel = {};
  for (var l = 0; l < levels.length; l++) {
    var levelNo = levels[l];
    var sequence = getLevelSetSequence_(levelNo) || [];
    var index = 0;
    var rounds = 0;

    if (sequence.length) {
      for (var j = 0; j < completedLogs.length; j++) {
        var logSetId = normalizeLevelPlainSetId_(completedLogs[j] && completedLogs[j].setId);
        if (!logSetId || logSetId.indexOf(String(levelNo) + '-') !== 0) continue;
        if (logSetId !== sequence[index]) continue;

        index += 1;
        if (index >= sequence.length) {
          rounds += 1;
          index = 0;
        }
      }
    }

    roundsByLevel[levelNo] = rounds;
  }

  var studiedSetCountByLevel = {};
  for (var hs = 0; hs < historyLogs.length; hs++) {
    var historyPlainSetId = normalizeLevelPlainSetId_(
      (historyLogs[hs] && (historyLogs[hs].setId || historyLogs[hs].Set_ID)) || ''
    );
    if (!historyPlainSetId) continue;
    var historyLevelNo = Number(String(historyPlainSetId).split('-')[0] || 0);
    if (!(historyLevelNo >= 3 && historyLevelNo <= 13)) continue;
    if (!studiedSetCountByLevel[historyLevelNo]) studiedSetCountByLevel[historyLevelNo] = {};
    studiedSetCountByLevel[historyLevelNo][historyPlainSetId] = true;
  }

  function wmGetStudiedSetCountForLevel__(levelNo) {
    var map = studiedSetCountByLevel[levelNo] || {};
    return Object.keys(map).length;
  }

  function wmGetDisplayedRoundForLevel__(levelNo) {
    var completedRound = Number(roundsByLevel[levelNo] || 0);
    var studiedCount = wmGetStudiedSetCountForLevel__(levelNo);
    if (studiedCount <= 0) return 0;
    if (Number(levelNo) === Number(currentLevel)) {
      return Math.max(1, completedRound + 1);
    }
    return Math.max(1, completedRound);
  }

  /* 순차완주세트는 학습한 세트 수가 아니라 현재 회차에서 순서대로 끝낸 세트 수입니다.
     실제 저장 직전의 정확한 current round/index 값은 wmBuildCurrentProgressCacheFields_()에서 최종 보정합니다. */
  result.순차완주세트 = currentLevel ? wmGetStudiedSetCountForLevel__(currentLevel) : 0;
  result.순차완주회차 = currentLevel ? wmGetDisplayedRoundForLevel__(currentLevel) : 0;
  result.히스토리레벨 = levels.join('|');
  result.히스토리횟수 = levels.map(function(level){ return wmGetDisplayedRoundForLevel__(level); }).join('|');

  /* 기존 Map.html/응답 호환 별칭입니다. 실제 신규 현재진행_DB 헤더에는 쓰이지 않습니다. */
  result.완료된레벨 = result.히스토리레벨;
  result.완료횟수 = result.히스토리횟수;
  return result;
}

function wmBuildCurrentProgressCacheFields_(ss, record, stepState, prefetchedRecordValues) {
  /* WM_CURRENT_PROGRESS_RIGHT_COLUMNS_FILL_20260624_V1
   * 8.현재진행_DB 오른쪽 계산 컬럼을 한 번에 채웁니다.
   * - 현재 위치: 기록Set_ID 기준
   * - 완료/세트진행률/레벨진행률/최근30건: 2.학습기록_DB 기준
   * - 학습모드: 1.학생관리_DB 기준
   */
  record = record || {};
  stepState = stepState || wmBuildCurrentProgressStepStateFromLearningRecord_(record);

  var studentId = String(record['학생ID'] || '').trim().toUpperCase();
  var learningRecordSetId = String(stepState.Set_ID || record['Set_ID'] || '').trim();
  var currentRecordSetId = String(stepState.기록Set_ID || learningRecordSetId || '').trim();
  var currentPlainSetId = normalizeLevelPlainSetId_(currentRecordSetId);
  var learningPlainSetId = normalizeLevelPlainSetId_(learningRecordSetId);
  var currentLevel = currentPlainSetId ? Number(currentPlainSetId.split('-')[0] || 0) : 0;
  var currentSetText = currentPlainSetId || '';
  var levelText = currentLevel ? (currentLevel + '레벨') : '';

  var learningMode = getStudentLearningMode_(studentId, '') || buildDefaultLearningMode_();
  var goal = buildLevelTargetRoundGoalFromLearningMode_(learningMode);
  var targetRounds = Number(goal.목표회차 || 1);
  var sequence = currentLevel ? getLevelSetSequence_(currentLevel) : [];
  var totalSetCount = sequence.length;

  var completedCountBySet = {};
  var completedLogs = [];
  var recentItems = [];
  var completeRoundForLearningSet = 0;

  try {
    var recordSheet = ss.getSheetByName('2.학습기록_DB');
    if (recordSheet) {
      var values = prefetchedRecordValues && prefetchedRecordValues.length
        ? prefetchedRecordValues
        : recordSheet.getDataRange().getDisplayValues();
      if (values && values.length > 1) {
        var headers = values[0].map(function(h) { return String(h || '').trim(); });
        var idxStudent = headers.indexOf('학생ID');
        var idxSetId = headers.indexOf('Set_ID');
        var idxDate = headers.indexOf('학습날짜');
        var idxRecordId = headers.indexOf('학습기록ID');
        var idxStatus = headers.indexOf('완료상태');
        var idxScore = headers.indexOf('점수');
        var idxCompleteRound = headers.indexOf('완료');

        for (var i = 1; i < values.length; i++) {
          var row = values[i] || [];
          var rowStudentId = idxStudent >= 0 ? String(row[idxStudent] || '').trim().toUpperCase() : '';
          if (rowStudentId !== studentId) continue;

          var rowSetId = idxSetId >= 0 ? String(row[idxSetId] || '').trim() : '';
          var rowPlainSetId = normalizeLevelPlainSetId_(rowSetId);
          if (!rowPlainSetId) continue;

          var rowDate = idxDate >= 0 ? String(row[idxDate] || '').trim() : '';
          var rowStatus = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : '';
          var rowScore = idxScore >= 0 ? String(row[idxScore] || '').trim() : '';
          var rowRecordId = idxRecordId >= 0 ? String(row[idxRecordId] || '').trim() : '';
          var rowCompleteRound = idxCompleteRound >= 0 ? String(row[idxCompleteRound] || '').trim() : '';
          var sortTime = getLearningRecordSortTime_(rowDate || '');

          recentItems.push({
            학습날짜: rowDate,
            학습기록ID: rowRecordId,
            Set_ID: rowSetId,
            완료상태: rowStatus,
            완료: rowCompleteRound,
            점수: rowScore,
            sortTime: sortTime,
            rowIndex: i
          });

          if (wmSafeIsCompleteStatusForCurrentProgress_(rowStatus)) {
            completedCountBySet[rowPlainSetId] = Number(completedCountBySet[rowPlainSetId] || 0) + 1;

            completedLogs.push({
              setId: rowPlainSetId,
              learningDate: rowDate,
              sortTime: sortTime,
              rowIndex: i
            });

            if (learningPlainSetId && rowPlainSetId === learningPlainSetId) {
              completeRoundForLearningSet += 1;
            }
          }
        }
      }
    }
  } catch (err) {}

  /* WM_CURRENT_PROGRESS_INCLUDE_INCOMING_RECORD_20260704_V1
   * 현재진행_DB 캐시는 2.학습기록_DB를 기준으로 만들지만,
   * 저장 직후 시트 읽기 반영이 늦는 경우가 있어 방금 저장한 record를
   * 최근30건/완료세트 계산에 보강합니다.
   */
  try {
    var incomingRecordId = String(record['학습기록ID'] || '').trim();
    var incomingSetId = String(record['Set_ID'] || '').trim();
    var incomingPlainSetId = normalizeLevelPlainSetId_(incomingSetId);
    var incomingExists = false;
    for (var ir = 0; ir < recentItems.length; ir++) {
      var recentRecordId = String(recentItems[ir] && recentItems[ir].학습기록ID || '').trim();
      if (incomingRecordId && recentRecordId === incomingRecordId) {
        incomingExists = true;
        break;
      }
    }

    if (!incomingExists && incomingPlainSetId) {
      var incomingDate = String(record['학습날짜'] || '').trim();
      var incomingStatus = String(record['완료상태'] || '').trim();
      var incomingScore = String(record['점수'] || '').trim();
      var incomingRound = String(record['완료'] || '').trim();
      var incomingSortTime = getLearningRecordSortTime_(incomingDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));

      recentItems.push({
        학습날짜: incomingDate,
        학습기록ID: incomingRecordId,
        Set_ID: incomingSetId,
        완료상태: incomingStatus,
        완료: incomingRound,
        점수: incomingScore,
        sortTime: incomingSortTime,
        rowIndex: 999999
      });

      if (wmSafeIsCompleteStatusForCurrentProgress_(incomingStatus)) {
        completedCountBySet[incomingPlainSetId] = Math.max(1, Number(completedCountBySet[incomingPlainSetId] || 0));
        completedLogs.push({
          setId: incomingPlainSetId,
          learningDate: incomingDate,
          sortTime: incomingSortTime,
          rowIndex: 999999
        });
        if (learningPlainSetId && incomingPlainSetId === learningPlainSetId) {
          completeRoundForLearningSet = Math.max(completeRoundForLearningSet, wmParseOfficialCompleteRound_(incomingRound) || 1);
        }
      }
    }
  } catch (incomingRecordCacheErr) {}

  recentItems.sort(function(a, b) {
    var at = Number(a && a.sortTime || 0);
    var bt = Number(b && b.sortTime || 0);
    if (at !== bt) return bt - at;
    return Number(b && b.rowIndex || 0) - Number(a && a.rowIndex || 0);
  });

  var compactRecent30 = recentItems.slice(0, 30).map(function(item) {
    return {
      학습날짜: item.학습날짜 || '',
      학습기록ID: item.학습기록ID || '',
      Set_ID: item.Set_ID || '',
      완료상태: item.완료상태 || '',
      완료: item.완료 || '',
      점수: item.점수 || ''
    };
  });

  var completedSetCount = 0;
  for (var s = 0; s < sequence.length; s++) {
    if (Number(completedCountBySet[sequence[s]] || 0) > 0) completedSetCount += 1;
  }

  var levelProgressPercent = totalSetCount > 0
    ? Math.round((completedSetCount / totalSetCount) * 100)
    : 0;
  var setProgressPercent = wmCalculateSetProgressPercentFromStepState_(stepState, record);

  completedLogs.sort(function(a, b) {
    var at = Number(a && a.sortTime || 0);
    var bt = Number(b && b.sortTime || 0);
    if (at !== bt) return at - bt;
    return Number(a && a.rowIndex || 0) - Number(b && b.rowIndex || 0);
  });

  var officialRound = 1;
  var officialIndex = 0;
  var officialCompletedRounds = 0;
  if (sequence.length) {
    for (var l = 0; l < completedLogs.length; l++) {
      if (officialRound > targetRounds) break;
      var logSetId = normalizeLevelPlainSetId_(completedLogs[l] && completedLogs[l].setId);
      if (!logSetId || logSetId.indexOf(String(currentLevel) + '-') !== 0) continue;
      if (logSetId !== sequence[officialIndex]) continue;

      officialIndex += 1;
      if (officialIndex >= sequence.length) {
        officialCompletedRounds += 1;
        officialRound += 1;
        officialIndex = 0;
      }
    }
  }

  if (officialCompletedRounds > targetRounds) officialCompletedRounds = targetRounds;

  /* WM_LEVEL_COMPLETE_COUNT_AXIS_20260707_V1
   * 레벨완료조건 = 교사가 정한 목표 완주 회차입니다.
   * 레벨완료횟수 = 학습기록_DB 기준 해당 레벨 전체 세트를 처음부터 끝까지 완주한 횟수입니다.
   * 순차완주회차 = 레벨완료횟수 + 1, 단 목표 회차 도달 후에는 레벨완료조건으로 고정합니다.
   */
  /* WM_LEVEL_COMPLETE_CONDITION_COUNT_FINAL_LOCK_20260727_V1
   * 레벨완료횟수는 공식 전환에서 확정한 현재진행_DB 최종값을 사용합니다.
   * 학습기록 재계산값이 이 값을 다시 덮어쓰지 못하게 합니다. */
  var lockedTransitionCount = record['__WM_LOCKED_LEVEL_COMPLETE_COUNT'];
  var storedLevelState = wmGetLockedLevelCompletionState_(studentId, currentLevel);
  var levelCompleteCount = lockedTransitionCount !== undefined && lockedTransitionCount !== ''
    ? wmNormalizeCurrentProgressRoundValue_(lockedTransitionCount)
    : wmNormalizeCurrentProgressRoundValue_(storedLevelState.levelCompleteCount);
  var levelComplete = totalSetCount > 0 && levelCompleteCount >= targetRounds;
  var nextOfficialSet = levelComplete ? getNextLevelFirstSetId_(currentLevel) : (sequence[officialIndex] || currentSetText || '');
  var levelHistoryCache = wmBuildCurrentProgressLevelHistoryCache_(completedLogs, currentLevel, recentItems);

  /* WM_CURRENT_PROGRESS_SEQUENTIAL_SET_COUNT_LOCK_20260707_V1
   * 순차완주세트 = 현재 회차에서 앞에서부터 순서대로 완료한 세트 수입니다.
   * - 현재 학습중 세트는 포함하지 않습니다.
   * - 완료세트수는 순서 무관 완료 총수, 순차완주세트는 공식 색상/레벨완료 판단용입니다.
   * 예: 5레벨 23세트 중 1~11 완료, 12세트 진행중이면 완료세트수=11, 순차완주세트=11, 순차완주회차=1.
   */
  var hasCurrentLevelCompletedLog = completedLogs.some(function(log) {
    var logSetId = normalizeLevelPlainSetId_(log && log.setId);
    return logSetId && logSetId.indexOf(String(currentLevel) + '-') === 0;
  });
  var sequentialSetCount = levelComplete ? totalSetCount : Math.max(0, Math.min(totalSetCount, Number(officialIndex || 0)));
  var sequentialRound = 0;
  if (hasCurrentLevelCompletedLog || sequentialSetCount > 0 || levelCompleteCount > 0) {
    sequentialRound = levelComplete ? targetRounds : (levelCompleteCount + 1);
  }
  sequentialRound = wmNormalizeCurrentProgressRoundValue_(Math.min(targetRounds || 1, sequentialRound));

  return {
    현재레벨: levelText,
    현재세트: currentSetText,
    레벨완료조건: learningMode.레벨완료조건 || '',
    레벨완료횟수: levelCompleteCount,
    순차완주세트: wmNormalizeCurrentProgressSetCountValue_(sequentialSetCount),
    순차완주회차: sequentialRound,
    히스토리레벨: levelHistoryCache.히스토리레벨,
    히스토리횟수: wmNormalizeCurrentProgressHistoryRoundsValue_(levelHistoryCache.히스토리횟수),

    /* 기존 응답 호환 별칭: 신규 현재진행_DB 헤더에는 없으면 저장되지 않습니다. */
    완료회차: completeRoundForLearningSet,
    완료된레벨: levelHistoryCache.히스토리레벨,
    완료횟수: levelHistoryCache.히스토리횟수,
    레벨회차추가: learningMode.레벨회차추가 || 0,
    완료세트수: completedSetCount,
    전체세트수: totalSetCount,
    세트진행률: wmFormatCurrentProgressPercentText_(setProgressPercent),
    레벨진행률: wmFormatCurrentProgressPercentText_(levelProgressPercent),
    진행률: wmFormatCurrentProgressPercentText_(levelProgressPercent),
    최근점수: wmExtractScoreForCurrentProgress_(record),
    S4실루엣단계: learningMode.S4실루엣단계 || '',
    S6테스트모드: learningMode.S6테스트모드 || '',
    레벨완료여부: levelComplete ? 'Y' : 'N',
    다음세트: nextOfficialSet || '',
    최근30건JSON: wmBuildRecent30JsonForCurrentProgress_(compactRecent30),
    캐시버전: 'v1',
    비고: ''
  };
}

function wmUpsertCurrentProgress_(ss, headers, record, prefetchedRecordValues) {
  /* WM_CURRENT_PROGRESS_CACHE_FROM_LEARNING_RECORD_V4_20260624
   * 8.현재진행_DB는 2.학습기록_DB를 원본으로 하는 최신 캐시입니다.
   * 학생ID 기준 기존 행은 모두 삭제하고, 학생당 1행만 2행에 저장합니다.
   * 핵심 공식: 학습기록 Set_ID 유지 + 완료Step → 현재Step → 기록Set_ID.
   */
  var sheet = ss.getSheetByName('8.현재진행_DB');
  if (!sheet) return;

  var values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length === 0) return;

  var progressHeaders = values[0].map(function(h) {
    return String(h || '').trim();
  });

  var idxStudent = progressHeaders.indexOf('학생ID');
  var studentId = String(record['학생ID'] || '').trim().toUpperCase();
  if (!studentId || idxStudent < 0) return;

  /* WM_CURRENT_PROGRESS_PROFILE_MERGE_20260704_V1
   * Study payload에 학교/학년/반명/교사명이 비어 있어도 현재진행_DB에는
   * 1.학생관리_DB 기준 기본정보를 반드시 보강합니다.
   */
  try {
    var profileForProgress = getStudentBasicInfoForMap_(studentId);
    if (profileForProgress) {
      record['학생이름'] = String(record['학생이름'] || profileForProgress.studentName || studentId || '').trim();
      record['학교'] = String(record['학교'] || profileForProgress.school || '').trim();
      record['학년'] = String(record['학년'] || profileForProgress.grade || '').trim();
      record['Class'] = String(record['Class'] || profileForProgress.className || '').trim();
      record['교사명'] = String(record['교사명'] || profileForProgress.teacherName || '').trim();
    }
  } catch (profileForProgressErr) {}

  var existingRecentScore = '';
  var existingRoundScores = {'1회차점수':'','2회차점수':'','3회차점수':''};
  var existingProgressSetId = '';
  var existingLearningMapScoreJson = '';
  var idxRecentScore = progressHeaders.indexOf('최근점수');
  var idxLearningMapScoreJson = progressHeaders.indexOf('학습맵점수JSON');
  var idxExistingSetId = progressHeaders.indexOf('Set_ID');
  var roundScoreIndexes = {
    '1회차점수': progressHeaders.indexOf('1회차점수'),
    '2회차점수': progressHeaders.indexOf('2회차점수'),
    '3회차점수': progressHeaders.indexOf('3회차점수')
  };
  for (var er = 1; er < values.length; er++) {
    var existingStudentId = String(values[er][idxStudent] || '').trim().toUpperCase();
    if (existingStudentId !== studentId) continue;
    if (idxRecentScore >= 0) existingRecentScore = String(values[er][idxRecentScore] || '').trim();
    if (idxLearningMapScoreJson >= 0) existingLearningMapScoreJson = String(values[er][idxLearningMapScoreJson] || '').trim();
    if (idxExistingSetId >= 0) existingProgressSetId = String(values[er][idxExistingSetId] || '').trim();
    Object.keys(roundScoreIndexes).forEach(function(headerName) {
      var idx = roundScoreIndexes[headerName];
      if (idx >= 0) existingRoundScores[headerName] = String(values[er][idx] || '').trim();
    });
    break;
  }

  var oldScoreLevel = parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(existingProgressSetId || ''));
  var newScoreLevel = parseLearningMapPlainSetId_(normalizeLevelPlainSetId_(record['Set_ID'] || ''));
  if (oldScoreLevel && newScoreLevel && oldScoreLevel.level !== newScoreLevel.level) {
    existingRoundScores = {'1회차점수':'','2회차점수':'','3회차점수':''};
  }

  var stepState = wmBuildCurrentProgressStepStateFromLearningRecord_(record);
  stepState.기록Set_ID = wmNormalizeCurrentProgressRecordSetId_(stepState.기록Set_ID, stepState.Set_ID || record['Set_ID'] || '');
  if (!stepState.기록Set_ID) return;
  var cacheFields = wmBuildCurrentProgressCacheFields_(ss, record, stepState, prefetchedRecordValues);
  if (cacheFields) {
    var incomingScoreForRound = wmExtractScoreForCurrentProgress_(record);
    var incomingRoundForScore = Number(cacheFields['완료회차'] || 0);
    var incomingRoundScoreHeader = wmCurrentProgressRoundScoreHeader_(incomingRoundForScore);
    if (stepState.완료Step === 'TEST' && incomingScoreForRound && incomingRoundScoreHeader) {
      existingRoundScores[incomingRoundScoreHeader] = incomingScoreForRound;
      cacheFields['최근점수'] = incomingScoreForRound;
      cacheFields['학습맵점수JSON'] = wmUpdateLearningMapScoreJson_(
        existingLearningMapScoreJson,
        record['Set_ID'],
        incomingRoundForScore,
        incomingScoreForRound
      );
    } else {
      cacheFields['최근점수'] = incomingScoreForRound || existingRecentScore || '';
      cacheFields['학습맵점수JSON'] = existingLearningMapScoreJson || '';
    }
    cacheFields['레벨완료횟수'] = wmNormalizeCurrentProgressRoundValue_(cacheFields['레벨완료횟수']);
    cacheFields['순차완주세트'] = wmNormalizeCurrentProgressSetCountValue_(cacheFields['순차완주세트']);
    cacheFields['순차완주회차'] = wmNormalizeCurrentProgressRoundValue_(cacheFields['순차완주회차']);
    cacheFields['히스토리횟수'] = wmNormalizeCurrentProgressHistoryRoundsValue_(cacheFields['히스토리횟수']);
  }

  var row = progressHeaders.map(function(header) {
    if (header === '학생ID') return record['학생ID'] || '';
    if (header === '학생이름') return record['학생이름'] || '';
    if (header === '학교') return record['학교'] || '';
    if (header === '학년') return record['학년'] || '';
    if (header === 'Class') return record['Class'] || '';
    if (header === '교사명') return record['교사명'] || '';

    if (header === 'Set_ID') return stepState.Set_ID || record['Set_ID'] || '';
    if (header === '완료Step') return stepState.완료Step || '';
    if (header === '현재Step') return stepState.현재Step || 'STEP1';
    if (header === '기록Set_ID') return stepState.기록Set_ID || '';

    if (header === '최종수정일') return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
    if (header === '학습기록ID') return record['학습기록ID'] || '';

    if (header === '레벨완료횟수') return cacheFields ? String(wmNormalizeCurrentProgressRoundValue_(cacheFields['레벨완료횟수'])) : '0';
    if (header === '순차완주세트') return cacheFields ? wmNormalizeCurrentProgressSetCountValue_(cacheFields['순차완주세트']) : 0;
    if (header === '순차완주회차') return cacheFields ? String(wmNormalizeCurrentProgressRoundValue_(cacheFields['순차완주회차'])) : '0';
    if (header === '1회차점수' || header === '2회차점수' || header === '3회차점수') return existingRoundScores[header] || '';
    if (header === '학습맵점수JSON') return cacheFields ? (cacheFields['학습맵점수JSON'] || existingLearningMapScoreJson || '') : (existingLearningMapScoreJson || '');
    if (header === '최근점수') return cacheFields ? wmExtractScoreForCurrentProgress_({
      '점수': cacheFields['최근점수'] || record['점수'],
      '한영주관식': record['한영주관식'],
      '최근점수': existingRecentScore
    }) : '';

    if (cacheFields && cacheFields[header] !== undefined) return cacheFields[header];

    if (header === 'Step1_총시간') return record['Step1_총시간'] || '';
    if (header === 'Step2_총시간') return record['Step2_총시간'] || '';
    if (header === 'Step3_총시간') return record['Step3_총시간'] || '';
    if (header === 'Step4_총시간') return record['Step4_총시간'] || '';
    if (header === 'Step5_총시간') return record['Step5_총시간'] || record['시험전재학습'] || record['시험전학습'] || '';
    if (header === '시험전학습') return record['시험전학습'] || record['Step5_총시간'] || record['시험전재학습'] || '';
    if (header === '시험전재학습') return record['시험전재학습'] || record['Step5_총시간'] || record['시험전학습'] || '';
    if (header === 'Test_총시간') return record['Test_총시간'] || '';
    if (header === '점수') return record['점수'] || '';
    if (header === '완료상태') return record['완료상태'] || '';
    if (record[header] !== undefined) return record[header];
    return '';
  });

  row = wmForceCurrentProgressRowFormulaValues_(progressHeaders, row, cacheFields, record, existingRecentScore);

  var matchingProgressRows = [];
  for (var mr = 2; mr <= values.length; mr++) {
    if (String(values[mr - 1][idxStudent] || '').trim().toUpperCase() === studentId) matchingProgressRows.push(mr);
  }

  /* 같은 학생이 계속 학습 중이면 최신행은 이미 2행입니다.
   * 이때 deleteRow + insertRowBefore를 생략하고 2행을 한 번에 갱신합니다. */
  var canOverwriteLatestRow = matchingProgressRows.length === 1 && matchingProgressRows[0] === 2;

  if (!canOverwriteLatestRow) for (var r = values.length; r >= 2; r--) {
    var rowStudent = String(values[r - 1][idxStudent] || '').trim().toUpperCase();
    if (rowStudent === studentId) {
      sheet.deleteRow(r);
    }
  }

  if (!canOverwriteLatestRow) sheet.insertRowBefore(2);

  /* WM_CURRENT_PROGRESS_ROUND_COLUMNS_TEXT_FORMAT_20260707_V1
   * 시트 열 서식이 퍼센트로 남아 있어도 회차 컬럼은 100%/0%로 표시되지 않게 텍스트 서식을 강제합니다.
   */
  ['레벨완료횟수','순차완주회차'].forEach(function(headerName) {
    var idx = progressHeaders.indexOf(headerName);
    if (idx >= 0) {
      try { sheet.getRange(2, idx + 1, 1, 1).setNumberFormat('@'); } catch (formatErr) {}
      row[idx] = String(wmNormalizeCurrentProgressRoundValue_(row[idx]));
    }
  });

  sheet.getRange(2, 1, 1, progressHeaders.length).setValues([row]);
  return {
    success: true,
    nextSetId: cacheFields ? String(cacheFields['다음세트'] || '').trim() : '',
    overwriteLatestRow: canOverwriteLatestRow
  };
}

/* WM_CURRENT_PROGRESS_LATEST_TOP_SORT_260618_V1
 * 8.현재진행_DB는 저장/갱신 직후 최종수정일 기준으로 최신 행이 위에 오도록 정렬합니다.
 * 1행 헤더는 고정하고 2행부터 데이터만 내림차순 정렬합니다.
 */
function wmSortCurrentProgressSheetLatestFirst_(sheet, headers) {
  try {
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow <= 2 || lastColumn < 1) return;

    var progressHeaders = headers && headers.length
      ? headers.map(function(h) { return String(h || '').trim(); })
      : sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });

    var idxUpdated = progressHeaders.indexOf('최종수정일');
    if (idxUpdated < 0) idxUpdated = progressHeaders.indexOf('학습날짜');
    if (idxUpdated < 0 && lastColumn >= 10) idxUpdated = 9;
    if (idxUpdated < 0) return;

    SpreadsheetApp.flush();
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort([{ column: idxUpdated + 1, ascending: false }]);
  } catch (err) {}
}

function wmMergeExistingLearningRecordTimes_(sheet, headers, targetRow, record) {
  if (!sheet || !headers || targetRow <= 1 || !record) return record;
  var existing = {};
  /* WM_SAVE_EXISTING_ROW_SINGLE_READ_SPEED_V1_20260818
   * 기존 기록의 Step시간 승계 공식은 그대로 두고, 컬럼별 getRange 반복을 한 행 1회 읽기로 통합합니다. */
  var existingRow = sheet.getRange(targetRow, 1, 1, headers.length).getDisplayValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim();
    if (key) existing[key] = String(existingRow[i] || '').trim();
  }
  return wmMergeRecordStepTimes_(existing, record);
}

function wmMergeRecordStepTimes_(existingRecord, incomingRecord) {
  existingRecord = existingRecord || {};
  incomingRecord = incomingRecord || {};
  var merged = {};

  Object.keys(existingRecord).forEach(function(k) { merged[k] = existingRecord[k]; });
  Object.keys(incomingRecord).forEach(function(k) {
    var value = incomingRecord[k];
    if (value !== undefined && value !== null && String(value).trim() !== '') {
      merged[k] = value;
    }
  });

  var fields = ['Step1_총시간', 'Step2_총시간', 'Step3_총시간', 'Step4_총시간', 'Step5_총시간', '시험전재학습', '시험전학습', 'Test_총시간'];
  for (var i = 0; i < fields.length; i++) {
    var f = fields[i];
    if (!String(merged[f] || '').trim()) {
      merged[f] = String(existingRecord[f] || incomingRecord[f] || '').trim();
    }
  }

  if (!String(merged['Step5_총시간'] || '').trim()) merged['Step5_총시간'] = String(merged['시험전재학습'] || merged['시험전학습'] || '').trim();
  if (!String(merged['시험전재학습'] || '').trim() && String(merged['Step5_총시간'] || merged['시험전학습'] || '').trim()) {
    merged['시험전재학습'] = String(merged['Step5_총시간'] || merged['시험전학습'] || '').trim();
  }
  if (!String(merged['시험전학습'] || '').trim() && String(merged['Step5_총시간'] || merged['시험전재학습'] || '').trim()) {
    merged['시험전학습'] = String(merged['Step5_총시간'] || merged['시험전재학습'] || '').trim();
  }

  var totalSeconds = parseTimeTextToSeconds_(merged['Step1_총시간'])
    + parseTimeTextToSeconds_(merged['Step2_총시간'])
    + parseTimeTextToSeconds_(merged['Step3_총시간'])
    + parseTimeTextToSeconds_(merged['Step4_총시간'])
    + parseTimeTextToSeconds_(merged['Step5_총시간'] || merged['시험전재학습'] || merged['시험전학습'])
    + parseTimeTextToSeconds_(merged['Test_총시간']);

  if (totalSeconds > 0) {
    merged['총소요시간'] = formatSecondsForSheet(totalSeconds);
  }

  return merged;
}

function getStepElapsed(completedSteps, stepId) {
  return getStepElapsedFromNormalizedSteps_(normalizeCompletedStepElapsedSeconds_(completedSteps), stepId);
}

function getStepElapsedFromNormalizedSteps_(normalizedSteps, stepId) {
  normalizedSteps = Array.isArray(normalizedSteps) ? normalizedSteps : [];

  for (var i = 0; i < normalizedSteps.length; i++) {
    var step = normalizedSteps[i] || {};
    if (step.stepId === stepId) {
      var unitSeconds = Number(step.__unitElapsedSeconds || 0);
      var formatted = formatSecondsForSheet(unitSeconds);
      if (formatted) {
        return formatted;
      }
      return String(step.elapsedText || step.timeText || '').trim();
    }
  }

  return '';
}

function deriveLastCompletedStepFromCompletedSteps_(completedSteps) {
  if (!Array.isArray(completedSteps) || !completedSteps.length) {
    return '';
  }

  var order = [
    ['step1_flash_card', 'STEP1'],
    ['step2_choose_meaning', 'STEP2'],
    ['step3_choose_english', 'STEP3'],
    ['step4_memory', 'STEP4'],
    ['step6_last_flash_review', 'STEP5'],
    ['test_all', 'TEST']
  ];

  var found = '';
  for (var i = 0; i < order.length; i++) {
    var stepId = order[i][0];
    for (var j = 0; j < completedSteps.length; j++) {
      if ((completedSteps[j] || {}).stepId === stepId) {
        found = order[i][1];
      }
    }
  }

  return found;
}

function deriveCurrentProgressStepFromCompletedSteps_(completedSteps, status) {
  var normalizedStatus = normalizeLearningStatus_(status);
  if (normalizedStatus === '완료') {
    return 'COMPLETE';
  }

  var last = deriveLastCompletedStepFromCompletedSteps_(completedSteps);
  if (last === 'STEP1') return 'STEP2';
  if (last === 'STEP2') return 'STEP3';
  if (last === 'STEP3') return 'STEP4';
  if (last === 'STEP4') return 'STEP5';
  if (last === 'STEP5') return 'TEST';
  if (last === 'BEFORE_TEST') return 'TEST';
  if (last === 'TEST') return 'COMPLETE';
  return 'STEP1';
}


function isOfficialLearningRecordId_(recordId) {
  return /^R\d{7}$/.test(String(recordId || '').trim());
}

function generateMonthlyLearningRecordId_(sheet, headers, now) {
  var tz = Session.getScriptTimeZone();
  var prefix = 'R' + Utilities.formatDate(now, tz, 'yyMM');
  var idxRecordId = headers.indexOf('학습기록ID');
  var idxDate = headers.indexOf('학습날짜');

  if (idxRecordId < 0) {
    throw new Error('2.학습기록_DB에서 학습기록ID 컬럼을 찾을 수 없습니다.');
  }

  var lastRow = sheet.getLastRow();
  var maxSeq = 0;
  var sameMonthRowCount = 0;

  if (lastRow >= 2) {
    /* WM_RECORD_ID_MIN_COLUMNS_SPEED_V1_20260826
     * 순번 계산에 실제로 필요한 학습기록ID/학습날짜 열만 1회 읽습니다.
     * ID 공식·월별 순번 계산·결과값은 기존과 동일합니다. */
    var minIndex = idxDate >= 0 ? Math.min(idxRecordId, idxDate) : idxRecordId;
    var maxIndex = idxDate >= 0 ? Math.max(idxRecordId, idxDate) : idxRecordId;
    var rangeValues = sheet.getRange(2, minIndex + 1, lastRow - 1, maxIndex - minIndex + 1).getDisplayValues();

    for (var i = 0; i < rangeValues.length; i++) {
      var row = rangeValues[i] || [];
      var id = String(row[idxRecordId - minIndex] || '').trim();

      if (id.indexOf(prefix) === 0 && /^R\d{7}$/.test(id)) {
        var seq = Number(id.slice(5));
        if (isFinite(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }

      if (idxDate >= 0) {
        var datePrefix = getLearningRecordMonthPrefix_(row[idxDate - minIndex]);
        if (datePrefix === prefix) {
          sameMonthRowCount += 1;
        }
      }
    }
  }

  /* WM_RECORD_ID_LEGACY_COUNT_FIX_V1
   * 이미 저장된 구 임시ID 행도 같은 월 순번에 포함합니다.
   * 예: 2026-06 기존 구 임시ID 2건이 있으면 다음 신규 ID는 R2606003입니다.
   */
  var nextSeq = Math.max(maxSeq, sameMonthRowCount) + 1;
  if (nextSeq > 999) {
    throw new Error(prefix + ' 월 학습기록ID 순번이 999개를 초과했습니다.');
  }

  return prefix + String(nextSeq).padStart(3, '0');
}

function buildLearningMapRecordIdDisplayState_(headers, values) {
  var state = {
    officialMaxByPrefix: {},
    nextByPrefix: {}
  };

  var idxRecordId = headers.indexOf('학습기록ID');
  if (idxRecordId < 0 || !values || values.length < 2) {
    return state;
  }

  for (var i = 1; i < values.length; i++) {
    var id = String((values[i] || [])[idxRecordId] || '').trim();
    if (!/^R\d{7}$/.test(id)) {
      continue;
    }

    var prefix = id.slice(0, 5);
    var seq = Number(id.slice(5));
    if (isFinite(seq) && seq > (state.officialMaxByPrefix[prefix] || 0)) {
      state.officialMaxByPrefix[prefix] = seq;
    }
  }

  return state;
}

function normalizeLearningMapRecordIdForDisplay_(recordId, learningDate, state) {
  var raw = String(recordId || '').trim();
  if (isOfficialLearningRecordId_(raw)) {
    return raw;
  }

  var prefix = getLearningRecordMonthPrefix_(learningDate);
  if (!prefix) {
    return raw;
  }

  state = state || { officialMaxByPrefix: {}, nextByPrefix: {} };
  var next = state.nextByPrefix[prefix];
  if (!next) {
    next = (state.officialMaxByPrefix[prefix] || 0) + 1;
  }

  if (next > 999) {
    return raw;
  }

  state.nextByPrefix[prefix] = next + 1;
  return prefix + String(next).padStart(3, '0');
}

function getLearningRecordMonthPrefix_(value) {
  var text = String(value || '').trim();
  var match = text.match(/(\d{4})[-.\/년\s]+(\d{1,2})/);

  if (match) {
    return 'R' + String(match[1]).slice(-2) + String(Number(match[2])).padStart(2, '0');
  }

  try {
    var date = value instanceof Date ? value : new Date(text);
    if (!isNaN(date.getTime())) {
      return 'R' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyMM');
    }
  } catch (err) {}

  return '';
}

function normalizeDeviceInfo_(userAgent) {
  var ua = String(userAgent || '').toLowerCase();

  if (!ua) {
    return '';
  }

  if (/ipad|tablet|sm-t|galaxy tab|nexus 7|nexus 9|kindle|silk/.test(ua)) {
    return 'Tap';
  }

  if (/mobile|iphone|android|phone|ipod/.test(ua)) {
    return 'Mobile';
  }

  return 'PC';
}

function normalizeLearningStatus_(status) {
  var value = String(status || '').trim();

  if (value === 'completed') {
    return '완료';
  }

  /* WM_STATUS_BLANK_MID_STEP_V1
   * 설계규칙: 중간 저장/복귀용 저장은 완료상태를 표시하지 않습니다.
   * 완료상태 컬럼은 최종 Test 완료 시에만 '완료'로 기록합니다.
   */
  if (value === 'step_completed' || value === 'in_progress' || value === 'left_mid_step') {
    return '';
  }

  return value;
}

function normalizeLearningRecordHeaders_(headers) {
  /* WM_CODEGS_NO_PROGRESS_EXTRA_COLUMNS_FIX_V1
   * 설계규칙: 현재Step / 완료Step 컬럼을 추가하지 않습니다.
   * 복귀 판단은 기존 Step 시간 컬럼만 사용합니다.
   */
  return (headers || []).map(function(h){ return String(h || '').trim(); });
}

function findLearningRecordRow_(sheet, headers, recordId, studentId, setId) {
  /* WM_RECORD_SESSION_ID_ONLY_UPSERT_20260617_V3
   * UPDATE 기준은 학습기록ID(recordSessionId)만 사용합니다.
   * 학생ID+Set_ID 기준 덮어쓰기는 분리학습 기록을 한 행으로 합쳐버리므로 사용하지 않습니다. */
  var idxRecordId = headers.indexOf('학습기록ID');
  var lastRow = sheet.getLastRow();

  if (lastRow < 2 || idxRecordId < 0) {
    return -1;
  }

  var wantedRecordId = String(recordId || '').trim();
  if (!wantedRecordId || !isOfficialLearningRecordId_(wantedRecordId)) {
    return -1;
  }

  /* WM_RECORD_ID_TEXTFINDER_SPEED_V1_20260826
   * 학습기록ID 열 전체를 Apps Script 메모리로 가져오지 않고 exact TextFinder로 같은 행을 찾습니다.
   * UPDATE 기준은 기존과 동일하게 학습기록ID 하나뿐입니다. */
  var finder = sheet
    .getRange(2, idxRecordId + 1, lastRow - 1, 1)
    .createTextFinder(wantedRecordId)
    .matchEntireCell(true)
    .matchCase(true);
  var cell = finder.findNext();
  return cell ? cell.getRow() : -1;
}

function wmFindLatestIncompleteLearningRecordRowByStudentSet_(sheet, headers, studentId, setId, expectedRecordId) {
  /* WM_SPLIT_LEARNING_CUMULATIVE_TIME_SOURCE_20260617_V1
   * 분리학습 새 행 INSERT 시 이전 미완료 기록의 Step 시간을 승계하기 위한 최신 행을 찾습니다.
   * 완료 행은 새 회차/복습의 출발점일 수 있으므로 승계 대상에서 제외합니다. */
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var idxStudent = headers.indexOf('학생ID');
  var idxSet = headers.indexOf('Set_ID');
  var idxStatus = headers.indexOf('완료상태');
  var idxRecordId = headers.indexOf('학습기록ID');

  if (idxStudent < 0 || idxSet < 0 || idxRecordId < 0 || !studentId || !setId || !expectedRecordId) {
    return -1;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  var wantedStudent = String(studentId || '').trim().toUpperCase();
  var wantedSet = normalizeStudySetIdForCompare_(setId);
  var wantedRecordId = String(expectedRecordId || '').trim();
  var targetRow = -1;

  for (var r = 0; r < values.length; r++) {
    var row = values[r] || [];
    var rowStudent = String(row[idxStudent] || '').trim().toUpperCase();
    var rowSet = normalizeStudySetIdForCompare_(row[idxSet]);
    var rowRecordId = String(row[idxRecordId] || '').trim();

    if (rowStudent !== wantedStudent || rowSet !== wantedSet || rowRecordId !== wantedRecordId) {
      continue;
    }

    var rowStatus = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : '';
    if (isCompleteStatus(rowStatus)) {
      continue;
    }

    targetRow = r + 2;
    break;
  }

  return targetRow;
}

function wmMergeLatestIncompleteSetTimesForNewRecord_(sheet, headers, record) {
  /* WM_SPLIT_LEARNING_CUMULATIVE_TIME_MERGE_20260617_V1
   * 같은 세트를 종료 후 이어서 학습해 새 행으로 저장할 때,
   * 이전 미완료 행의 Step1~Step5/Test 시간을 현재 행에 승계합니다.
   * 총소요시간은 승계된 Step 시간 + 이번 저장의 Step/Test 시간 합산값으로 재계산됩니다.
   * Test_총시간은 총소요시간 안에 포함된 시험 구간만 별도 표시합니다. */
  if (!sheet || !headers || !record) return record;

  /* WM_REAL_ROUND_RECORD_GUARD_V1_20260727
   * 회차가 다른 과거 미완료 기록의 Step 시간을 새 회차에 합치지 않습니다.
   * 현재진행_DB가 가리키는 동일 학습기록ID만 이어학습 승계 대상으로 허용합니다. */
  var expectedRecordId = '';
  try {
    var progressSheet = sheet.getParent().getSheetByName('8.현재진행_DB');
    if (progressSheet && progressSheet.getLastRow() >= 2) {
      var progressValues = progressSheet.getDataRange().getDisplayValues();
      var progressHeaders = progressValues[0].map(function(h) { return String(h || '').trim(); });
      var progressStudentIndex = progressHeaders.indexOf('학생ID');
      var progressSetIndex = progressHeaders.indexOf('기록Set_ID');
      if (progressSetIndex < 0) progressSetIndex = progressHeaders.indexOf('Set_ID');
      var progressRecordIndex = progressHeaders.indexOf('학습기록ID');
      var wantedStudent = String(record['학생ID'] || '').trim().toUpperCase();
      var wantedSet = normalizeStudySetIdForCompare_(record['Set_ID']);

      if (progressStudentIndex >= 0 && progressSetIndex >= 0 && progressRecordIndex >= 0) {
        for (var pr = 1; pr < progressValues.length; pr++) {
          var progressStudent = String(progressValues[pr][progressStudentIndex] || '').trim().toUpperCase();
          var progressSet = normalizeStudySetIdForCompare_(progressValues[pr][progressSetIndex]);
          if (progressStudent === wantedStudent && progressSet === wantedSet) {
            expectedRecordId = String(progressValues[pr][progressRecordIndex] || '').trim();
            break;
          }
        }
      }
    }
  } catch (progressRecordGuardErr) {
    expectedRecordId = '';
  }

  if (!expectedRecordId || !isOfficialLearningRecordId_(expectedRecordId)) return record;

  var sourceRow = wmFindLatestIncompleteLearningRecordRowByStudentSet_(
    sheet,
    headers,
    record['학생ID'],
    record['Set_ID'],
    expectedRecordId
  );

  if (sourceRow <= 1) return record;

  return wmMergeExistingLearningRecordTimes_(sheet, headers, sourceRow, record);
}

function getLearningProgressData(studentId, setId) {
  var e = {
    parameter: {
      studentId: studentId,
      set_id: setId
    }
  };

  var result = buildLearningProgressResult_(e);
  return result;
}

function getLearningProgress(e) {
  try {
    return outputResult(e, buildLearningProgressResult_(e));
  } catch (err) {
    return outputResult(e, {
      success: false,
      message: '학습 진행상태 조회 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}

/* WM_STUDY_PROGRESS_FAST_V1_20260711
 * 8.현재진행_DB에서 학생 1행과 핵심 4개 필드만 직접 읽습니다.
 */
function getStudyProgressFast(e) {
  try {
    var studentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid)) || '').trim();
    var requestedSetId = String((e && e.parameter && (e.parameter.set_id || e.parameter.setId || e.parameter.Set_ID)) || '').trim().toUpperCase();
    var sessionToken = String((e && e.parameter && (e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN)) || '').trim();
    if (sessionToken && !wmIsStudentSessionValid_(studentId, sessionToken)) return outputResult(e, wmBuildSessionExpiredResponse_());
    if (!studentId) return outputResult(e, {success:false, message:'studentId가 없습니다.'});

    /* WM_STUDY_CURRENT_PROGRESS_DIRECT_ONE_ROW_V1_20260821
     * Study 진입 위치/학습조건은 8.현재진행_DB 학생 1행을 그대로 사용합니다.
     * 동일학생 findAll/최종수정일 비교/재계산/학생관리_DB 조회를 하지 않습니다. */
    var currentProgress = wmGetCurrentProgressCacheForStudent_(studentId);
    if (!currentProgress) {
      return outputResult(e, {
        success:false,
        sessionValid:sessionToken ? true : undefined,
        valid:sessionToken ? true : undefined,
        source:'8.현재진행_DB_DIRECT',
        studentId:studentId,
        currentProgressMissing:true,
        message:'8.현재진행_DB에서 학생 현재진행 1행을 찾을 수 없습니다.'
      });
    }

    var setId = String(currentProgress['Set_ID'] || '').trim().toUpperCase();
    var lastStep = String(currentProgress['완료Step'] || '').trim().toUpperCase();
    var currentStep = String(currentProgress['현재Step'] || '').trim().toUpperCase();
    var recordSetId = String(currentProgress['기록Set_ID'] || '').trim().toUpperCase();
    var learningRecordId = String(currentProgress['학습기록ID'] || '').trim();
    if (!recordSetId || !currentStep) {
      return outputResult(e, {
        success:false,
        sessionValid:sessionToken ? true : undefined,
        valid:sessionToken ? true : undefined,
        source:'8.현재진행_DB_DIRECT',
        studentId:studentId,
        setId:setId,
        recordSetId:recordSetId,
        currentProgressStep:currentStep,
        currentProgressInvalid:true,
        message:'8.현재진행_DB의 기록Set_ID/현재Step을 확인하세요.'
      });
    }

    var requestedMatchesProgress = !requestedSetId ||
      normalizeStudySetIdForCompare_(requestedSetId) === normalizeStudySetIdForCompare_(recordSetId);

    var headers = Object.keys(currentProgress).filter(function(key){ return String(key).indexOf('__') !== 0; });
    var core = headers.map(function(header){ return currentProgress[header] !== undefined ? currentProgress[header] : ''; });
    var progressLearningMode = requestedMatchesProgress ? buildLearningModeFromRow_(core, headers) : null;

    if (!requestedMatchesProgress) {
      return outputResult(e, {
        success:false,
        sessionValid:sessionToken ? true : undefined,
        valid:sessionToken ? true : undefined,
        source:'8.현재진행_DB_DIRECT',
        studentId:studentId,
        requestedSetId:requestedSetId,
        setId:setId,
        recordSetId:recordSetId,
        currentProgressStep:currentStep,
        learningRecordId:learningRecordId,
        setMismatch:true,
        message:'요청 세트와 8.현재진행_DB 기록Set_ID가 일치하지 않습니다.'
      });
    }

    return outputResult(e, {
      success:true,
      sessionValid:!!sessionToken,
      valid:!!sessionToken,
      message:'현재진행_DB 학생 1행 직접조회 성공',
      source:'8.현재진행_DB_DIRECT',
      studentId:studentId,
      setId:setId,
      recordSetId:recordSetId,
      learningRecordId:learningRecordId,
      currentProgressStep:currentStep,
      lastCompletedStep:lastStep,
      completedSteps:[],
      learningMode:progressLearningMode,
      modeSource:progressLearningMode ? '8.현재진행_DB_CURRENT_SET_LOCK' : '',
      currentProgress:currentProgress,
      currentProgressCache:currentProgress,
      currentProgressRow:currentProgress
    });
  } catch (err) {
    return outputResult(e, {success:false, message:'현재진행 직접조회 오류', error:String(err && err.message ? err.message : err)});
  }
}

function buildLearningProgressResult_(e) {
  var studentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid)) || '').trim();
  var setId = String((e && e.parameter && (e.parameter.set_id || e.parameter.setId || e.parameter.Set_ID || e.parameter.currentSetId)) || '').trim().toUpperCase();
  var sessionToken = String((e && e.parameter && (e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN)) || '').trim();

  if (sessionToken && !wmIsStudentSessionValid_(studentId, sessionToken)) {
    return wmBuildSessionExpiredResponse_();
  }

  if (!studentId || !setId) {
    return {
      success: false,
      message: 'studentId 또는 set_id가 없습니다.',
      studentId: studentId,
      setId: setId
    };
  }

  var ss = getLmsSpreadsheet_();
  if (!ss) {
    return {
      success: false,
      message: '스프레드시트 연결 실패',
      studentId: studentId,
      setId: setId
    };
  }

  var wantedStudent = studentId;
  var wantedSet = normalizeStudySetIdForCompare_(setId);

  /* 1순위: 8.현재진행_DB */
  var progressSheet = ss.getSheetByName('8.현재진행_DB');
  if (progressSheet) {
    var progressValues = progressSheet.getDataRange().getDisplayValues();

    if (progressValues && progressValues.length >= 2) {
      var progressHeaders = progressValues[0].map(function(h) {
        return String(h || '').trim();
      });

      var idxPStudent = progressHeaders.indexOf('학생ID');
      var idxPSet = progressHeaders.indexOf('Set_ID');
      var idxPRecordSet = progressHeaders.indexOf('기록Set_ID');
      var idxPDate = progressHeaders.indexOf('학습날짜');
      var idxPStatus = progressHeaders.indexOf('완료상태');
      var idxPCurrent = progressHeaders.indexOf('현재Step');
      var idxPLast = progressHeaders.indexOf('완료Step');

      if (idxPStudent >= 0 && (idxPSet >= 0 || idxPRecordSet >= 0)) {
        var latestProgressRecord = null;
        var latestProgressDate = '';

        for (var p = 1; p < progressValues.length; p++) {
          var prow = progressValues[p] || [];
          var pStudent = String(prow[idxPStudent] || '').trim().toUpperCase();
          var pSet = idxPRecordSet >= 0
            ? normalizeStudySetIdForCompare_(prow[idxPRecordSet])
            : '';
          if (!pSet && idxPSet >= 0) {
            pSet = normalizeStudySetIdForCompare_(prow[idxPSet]);
          }

          if (pStudent !== wantedStudent || pSet !== wantedSet) {
            continue;
          }

          var pDate = idxPDate >= 0 ? String(prow[idxPDate] || '') : '';
          if (!latestProgressRecord || pDate >= latestProgressDate) {
            latestProgressRecord = buildRecordObjectByHeaders_(progressHeaders, prow);
            latestProgressDate = pDate;
          }
        }

        if (latestProgressRecord) {
          var progress = deriveProgressFromLearningRecord_(latestProgressRecord);

          if (idxPCurrent >= 0 && String(latestProgressRecord['현재Step'] || '').trim()) {
            progress.currentProgressStep = String(latestProgressRecord['현재Step'] || '').trim();
          }

          if (idxPLast >= 0 && String(latestProgressRecord['완료Step'] || '').trim()) {
            progress.lastCompletedStep = String(latestProgressRecord['완료Step'] || '').trim();
          }

          return {
            success: true,
            message: '현재진행 조회 성공',
            source: '8.현재진행_DB',
            studentId: studentId,
            setId: setId,
            currentProgressStep: progress.currentProgressStep || 'STEP1',
            lastCompletedStep: progress.lastCompletedStep || '',
            completedSteps: progress.completedSteps || [],
            record: latestProgressRecord
          };
        }
      }
    }
  }

  /* 2순위: 기존 2.학습기록_DB fallback */
  var sheet = ss.getSheetByName('2.학습기록_DB');

  if (!sheet) {
    return {
      success: false,
      message: '2.학습기록_DB 시트를 찾을 수 없습니다.',
      studentId: studentId,
      setId: setId
    };
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) {
    return {
      success: true,
      message: '저장된 학습 진행상태가 없습니다.',
      source: 'none',
      studentId: studentId,
      setId: setId,
      currentProgressStep: 'STEP1',
      lastCompletedStep: '',
      completedSteps: [],
      record: null
    };
  }

  var headers = values[0].map(function(h) {
    return String(h || '').trim();
  });

  var idxStudent = headers.indexOf('학생ID');
  var idxSet = headers.indexOf('Set_ID');
  var idxDate = headers.indexOf('학습날짜');

  if (idxStudent < 0 || idxSet < 0) {
    return {
      success: false,
      message: '학생ID 또는 Set_ID 헤더를 찾을 수 없습니다.',
      studentId: studentId,
      setId: setId
    };
  }

  var latestRecord = null;
  var latestDate = '';

  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    var rowStudent = String(row[idxStudent] || '').trim().toUpperCase();
    var rowSet = normalizeStudySetIdForCompare_(row[idxSet]);

    if (rowStudent !== wantedStudent || rowSet !== wantedSet) {
      continue;
    }

    var rowDate = idxDate >= 0 ? String(row[idxDate] || '') : '';
    if (!latestRecord || rowDate >= latestDate) {
      latestRecord = buildRecordObjectByHeaders_(headers, row);
      latestDate = rowDate;
    }
  }

  if (!latestRecord) {
    return {
      success: true,
      message: '저장된 학습 진행상태가 없습니다.',
      source: '2.학습기록_DB',
      studentId: studentId,
      setId: setId,
      currentProgressStep: 'STEP1',
      lastCompletedStep: '',
      completedSteps: [],
      record: null
    };
  }

  var fallbackProgress = deriveProgressFromLearningRecord_(latestRecord);

  return {
    success: true,
    message: '학습 진행상태 조회 성공',
    source: '2.학습기록_DB',
    studentId: studentId,
    setId: setId,
    currentProgressStep: fallbackProgress.currentProgressStep || 'STEP1',
    lastCompletedStep: fallbackProgress.lastCompletedStep || '',
    completedSteps: fallbackProgress.completedSteps || [],
    record: latestRecord
  };
}

function normalizeStudySetIdForCompare_(setId) {
  return String(setId || '').trim().toUpperCase().replace(/^WM/, '');
}

function deriveProgressFromLearningRecord_(record) {
  record = record || {};

  var status = String(record['완료상태'] || '').trim();
  var completedSteps = [];

  function pushIfExists(columnName, stepId, stepTitle) {
    var elapsedText = String(record[columnName] || '').trim();
    if (!elapsedText) return;
    completedSteps.push({
      stepId: stepId,
      stepTitle: stepTitle,
      completedAt: String(record['학습날짜'] || ''),
      elapsedSeconds: 0,
      elapsedText: elapsedText
    });
  }

  pushIfExists('Step1_총시간', 'step1_flash_card', 'Step 1 Flash Card');
  pushIfExists('Step2_총시간', 'step2_choose_meaning', 'Step 2 Choose meaning');
  pushIfExists('Step3_총시간', 'step3_choose_english', 'Step 3 Choose English');
  pushIfExists('Step4_총시간', 'step4_memory', 'Step 4 Memory');
  pushIfExists('Step5_총시간', 'step6_last_flash_review', 'Step 5 Last Flash Review');
  pushIfExists('시험전학습', 'step6_last_flash_review', 'Step 5 Last Flash Review');
  pushIfExists('시험전재학습', 'step6_last_flash_review', 'Step 5 Last Flash Review');

  var last = deriveLastCompletedStepFromCompletedSteps_(completedSteps);
  var current = '';

  if (status === '완료') {
    current = 'COMPLETE';
  } else if (last === 'STEP1') {
    current = 'STEP2';
  } else if (last === 'STEP2') {
    current = 'STEP3';
  } else if (last === 'STEP3') {
    current = 'STEP4';
  } else if (last === 'STEP4') {
    current = 'STEP5';
  } else if (last === 'STEP5' || last === 'BEFORE_TEST') {
    current = 'TEST';
  } else {
    current = 'STEP1';
  }

  return {
    currentProgressStep: current,
    lastCompletedStep: last,
    completedSteps: completedSteps
  };
}

/* WM_DEBUG_CHECK_MODE_V1
 * /exec?mode=wmDebugCheck&studentId=S4821&setId=WM4-1-1
 * Code.gs가 실제로 현재 배포에 반영되었는지, Map/Study 화면을 거치지 않고 서버 응답만 바로 확인합니다.
 * 운영 로직은 바꾸지 않고 진단 JSON만 반환합니다.
 */
function wmDebugCheck(e) {
  try {
    var studentId = String((e && e.parameter && (e.parameter.studentId || e.parameter.studentID || e.parameter.sid)) || '').trim().toUpperCase();
    var setIdRaw = String((e && e.parameter && (e.parameter.setId || e.parameter.set_id || e.parameter.Set_ID || e.parameter.currentSetId)) || '').trim().toUpperCase();
    var setId = normalizeStudySetIdForCompare_(setIdRaw || 'WM4-1-1');
    var wmSetId = 'WM' + setId;

    var mapPayload = studentId ? buildLearningMapPayload_(studentId) : {
      success: false,
      message: 'studentId가 없습니다.'
    };

    var records = mapPayload && Array.isArray(mapPayload.records) ? mapPayload.records : [];
    var setStatusMap = mapPayload && mapPayload.setStatusMap ? mapPayload.setStatusMap : {};

    function statusFor(id) {
      var plain = normalizeStudySetIdForCompare_(id);
      return setStatusMap[id] || setStatusMap['WM' + plain] || setStatusMap[plain] || null;
    }

    var targetStatus = statusFor(setIdRaw || wmSetId);
    var check416 = statusFor('WM4-1-6');

    var targetRecords = [];
    for (var i = 0; i < records.length; i++) {
      var recordSetId = normalizeStudySetIdForCompare_(records[i] && records[i].Set_ID);
      if (recordSetId === setId) {
        targetRecords.push(records[i]);
      }
    }

    var sortedCheck = true;
    var sortSamples = [];
    for (var s = 0; s < Math.min(records.length, 10); s++) {
      sortSamples.push({
        학습날짜: records[s]['학습날짜'] || '',
        Set_ID: records[s]['Set_ID'] || '',
        점수: records[s]['점수'] || '',
        총소요시간: records[s]['총소요시간'] || '',
        한영주관식: records[s]['한영주관식'] || ''
      });
      if (s > 0 && getLearningRecordSortTime_(records[s - 1]['학습날짜']) < getLearningRecordSortTime_(records[s]['학습날짜'])) {
        sortedCheck = false;
      }
    }

    var latestTarget = targetRecords.length ? targetRecords[0] : null;
    var latestScoreCheck = latestTarget ? {
      savedScore: String(latestTarget['점수'] || ''),
      koEngWrite: String(latestTarget['한영주관식'] || ''),
      isScoreSameAsKoEngWrite: String(latestTarget['점수'] || '') === String(latestTarget['한영주관식'] || '')
    } : null;

    var debugResult = {
      success: true,
      debugMode: 'WM_DEBUG_CHECK_MODE_V1',
      message: 'wmDebugCheck 진단 mode가 현재 배포된 Code.gs에서 실행 중입니다.',
      requested: {
        studentId: studentId,
        setIdRaw: setIdRaw,
        normalizedSetId: wmSetId
      },
      codeMarkers: {
        totalTimeLogic: 'WM_TOTAL_TIME_STEP_UNIT_SUM_V5',
        stepTimeNormalizeLogic: 'WM_STEP_TIME_CUMULATIVE_TO_UNIT_V1 + normalized step columns' ,
        scoreLogic: '점수는 한영주관식 마지막 점수만 저장합니다. 평균점수는 사용하지 않습니다.',
        mapProgressLogic: 'setStatusMap.progressPercent / completedStepCount 기준 진단',
        recordSortLogic: 'records 최신순 정렬 진단'
      },
      mapPayloadSuccess: !!(mapPayload && mapPayload.success),
      mapPayloadMessage: mapPayload && mapPayload.message ? mapPayload.message : '',
      recordsCount: records.length,
      targetRecordsCount: targetRecords.length,
      targetSetStatus: targetStatus,
      check_4_1_6_status: check416,
      warnings: {
        fourOneSixHasProgress: !!(check416 && Number(check416.progressPercent || check416.progress || 0) > 0),
        recordsNotLatestFirst: !sortedCheck,
        latestScoreNotKoEngWrite: latestScoreCheck ? !latestScoreCheck.isScoreSameAsKoEngWrite : null
      },
      latestTargetRecord: latestTarget,
      latestScoreCheck: latestScoreCheck,
      first10RecordsSortSample: sortSamples,
      studentProfile: mapPayload && mapPayload.studentProfile ? mapPayload.studentProfile : null
    };

    return outputResult(e, debugResult);
  } catch (err) {
    return outputResult(e, {
      success: false,
      debugMode: 'WM_DEBUG_CHECK_MODE_V1',
      message: 'wmDebugCheck 서버 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}

function doPost(e) {
  return saveLearningRecord(e);
}


/* =========================================================
 * Word Mate DEV CENTER DB API
 * action=getSetData&set_id=WM4-1-1
 * 외부 콘텐츠/음원 DB에서 단어와 Audio_URL을 조회합니다.
 * ========================================================= */
function getSetData(e) {
  var setId = '';
  var studentId = '';
  var sessionToken = '';
  if (e && e.parameter) {
    setId = String(e.parameter.set_id || e.parameter.setId || e.parameter.Set_ID || '').trim();
    studentId = String(e.parameter.studentId || e.parameter.studentID || e.parameter.sid || '').trim();
    sessionToken = String(e.parameter.sessionToken || e.parameter.wmSessionToken || e.parameter.WM_SESSION_TOKEN || '').trim();
  }

  /* WM_REAL_SET_DATA_SESSION_DUPLICATE_RECOVERY_V1_20260812
   * Study 시작 직전 checkStudentSession이 별도로 실행됩니다.
   * 공개 학습 콘텐츠 조회에서 동일 학생관리_DB 세션을 다시 읽어
   * 콘텐츠·모드·진행 조회가 함께 밀리던 중복 검사를 제거합니다.
   * 학습기록 저장 API의 세션 검증은 그대로 유지합니다. */

  if (!setId) {
    return outputResult(e, {
      success: false,
      message: 'set_id가 없습니다.',
      set_id: setId,
      count: 0,
      words: []
    });
  }

  var parsed = parseWordMateSetId(setId);
  if (!parsed) {
    return outputResult(e, {
      success: false,
      message: 'Set_ID 형식이 올바르지 않습니다. 예: WM4-1-1',
      set_id: setId,
      count: 0,
      words: []
    });
  }

  var normalizedSetId = 'WM' + parsed.level + '-' + parsed.part + '-' + parsed.set;
  var cacheKey = wmCacheKey_('SET_DATA', normalizedSetId);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && cached.success) {
    cached.cacheHit = true;
    return outputResult(e, cached);
  }

  try {
    var contentSS = SpreadsheetApp.openById(WM_CONTENT_SPREADSHEET_ID);
    var levelSheetName = 'Level_' + pad2(parsed.level);
    var audioSheetName = 'L' + parsed.level;

    var levelSheet = contentSS.getSheetByName(levelSheetName);
    if (!levelSheet) {
      return outputResult(e, {
        success: false,
        message: levelSheetName + ' 시트를 찾을 수 없습니다.',
        set_id: normalizedSetId,
        sheet: levelSheetName,
        count: 0,
        words: []
      });
    }

    var audioInfo = getAudioInfoForSet(contentSS, audioSheetName, normalizedSetId);
    var words = getWordsForSet(levelSheet, normalizedSetId, audioInfo);
    var result = {
      success: true,
      message: 'Set 데이터 조회 성공',
      set_id: normalizedSetId,
      sheet: levelSheetName,
      audioSheet: audioSheetName,
      count: words.length,
      audio: audioInfo,
      words: words,
      cacheHit: false
    };

    wmCachePutJson_(cacheKey, result, 21600);
    return outputResult(e, result);

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: 'getSetData 서버 오류',
      error: String(err && err.message ? err.message : err),
      set_id: normalizedSetId,
      count: 0,
      words: []
    });
  }
}

function parseWordMateSetId(setId) {
  var text = String(setId || '').trim().toUpperCase();
  var match = text.match(/^WM\s*(\d+)\s*-\s*(\d+)\s*-\s*(\d+)$/);

  if (!match) {
    match = text.match(/^(\d+)\s*-\s*(\d+)\s*-\s*(\d+)$/);
  }

  if (!match) {
    return null;
  }

  return {
    level: Number(match[1]),
    part: Number(match[2]),
    set: Number(match[3])
  };
}

function pad2(num) {
  var n = Number(num);
  return n < 10 ? '0' + n : String(n);
}

function normalizeHeaderName(value) {
  return String(value || '').trim();
}

function buildHeaderMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeHeaderName(headers[i]);
    if (key && map[key] === undefined) {
      map[key] = i;
    }
  }
  return map;
}

function getCell(row, headerMap, keys) {
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (headerMap[key] !== undefined && headerMap[key] >= 0) {
      return String(row[headerMap[key]] || '').trim();
    }
  }
  return '';
}

function wmGetSheetHeaderMapFast_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (!lastCol) return { headers: [], headerMap: {} };

  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  for (var h = 0; h < headers.length; h++) {
    headers[h] = normalizeHeaderName(headers[h]);
  }

  return {
    headers: headers,
    headerMap: buildHeaderMap(headers)
  };
}

function wmFindRowsBySetIdFast_(sheet, idxSetId, setId) {
  var lastRow = sheet.getLastRow();
  if (!lastRow || lastRow < 2 || idxSetId === undefined || idxSetId < 0) return [];

  var col = Number(idxSetId) + 1;
  var finder = sheet
    .getRange(2, col, lastRow - 1, 1)
    .createTextFinder(String(setId || '').trim())
    .matchEntireCell(true)
    .matchCase(false);

  var ranges = finder.findAll() || [];
  var rows = [];
  for (var i = 0; i < ranges.length; i++) {
    rows.push(ranges[i].getRow());
  }
  rows.sort(function(a, b) { return a - b; });
  return rows;
}

function getAudioInfoForSet(contentSS, audioSheetName, setId) {
  var audioInfo = {
    Set_ID: setId,
    Drive_URL: '',
    Audio_URL: ''
  };

  var cacheKey = wmCacheKey_('AUDIO_INFO', setId);
  var cached = wmCacheGetJson_(cacheKey);
  if (cached && cached.Set_ID) return cached;

  var sheet = contentSS.getSheetByName(audioSheetName);
  if (!sheet) {
    return audioInfo;
  }

  var meta = wmGetSheetHeaderMapFast_(sheet);
  var headerMap = meta.headerMap || {};
  var idxSetId = headerMap['Set_ID'];

  if (idxSetId === undefined || idxSetId < 0) {
    return audioInfo;
  }

  var rows = wmFindRowsBySetIdFast_(sheet, idxSetId, setId);
  if (!rows.length) {
    wmCachePutJson_(cacheKey, audioInfo, 3600);
    return audioInfo;
  }

  var row = sheet.getRange(rows[0], 1, 1, sheet.getLastColumn()).getDisplayValues()[0] || [];
  audioInfo.Drive_URL = getCell(row, headerMap, ['Drive_URL', 'Drive URL', '드라이브URL']);
  audioInfo.Audio_URL = getCell(row, headerMap, ['Audio_URL', 'Audio URL', '음원URL']);

  wmCachePutJson_(cacheKey, audioInfo, 21600);
  return audioInfo;
}

function getWordsForSet(levelSheet, setId, audioInfo) {
  var cacheKey = wmCacheKey_('WORDS_ONLY', setId);
  var cachedWords = wmCacheGetJson_(cacheKey);
  if (cachedWords && cachedWords.length !== undefined) {
    for (var cw = 0; cw < cachedWords.length; cw++) {
      cachedWords[cw].Drive_URL = audioInfo && audioInfo.Drive_URL ? audioInfo.Drive_URL : '';
      cachedWords[cw].Audio_URL = audioInfo && audioInfo.Audio_URL ? audioInfo.Audio_URL : '';
    }
    return cachedWords;
  }

  var lastRow = levelSheet.getLastRow();
  var lastCol = levelSheet.getLastColumn();
  if (!lastRow || lastRow < 2 || !lastCol) {
    return [];
  }

  var meta = wmGetSheetHeaderMapFast_(levelSheet);
  var headerMap = meta.headerMap || {};
  var idxSetId = headerMap['Set_ID'];

  if (idxSetId === undefined || idxSetId < 0) {
    return [];
  }

  var targetRows = wmFindRowsBySetIdFast_(levelSheet, idxSetId, setId);
  if (!targetRows.length) {
    wmCachePutJson_(cacheKey, [], 1200);
    return [];
  }

  var minRow = targetRows[0];
  var maxRow = targetRows[targetRows.length - 1];
  var rowCount = maxRow - minRow + 1;
  var blockValues = levelSheet.getRange(minRow, 1, rowCount, lastCol).getDisplayValues();
  var targetRowMap = {};

  for (var tr = 0; tr < targetRows.length; tr++) {
    targetRowMap[targetRows[tr]] = true;
  }

  var words = [];

  for (var i = 0; i < blockValues.length; i++) {
    var actualRowNumber = minRow + i;
    if (!targetRowMap[actualRowNumber]) continue;

    var row = blockValues[i] || [];
    var rowSetId = String(row[idxSetId] || '').trim();
    if (rowSetId !== setId) continue;

    var item = {
      Level: getCell(row, headerMap, ['Level']),
      Part: getCell(row, headerMap, ['Part']),
      Set: getCell(row, headerMap, ['Set']),
      Set_ID: rowSetId,
      Word_No: getCell(row, headerMap, ['Word_No', 'Word No']),
      Word_ID: getCell(row, headerMap, ['Word_ID', 'Word ID']),
      Word: getCell(row, headerMap, ['Word']),
      POS: getCell(row, headerMap, ['POS']),
      Meaning_KR: getCell(row, headerMap, ['Meaning_KR', 'Meaning KR', '뜻']),
      Example_EN: getCell(row, headerMap, ['Example_EN', 'Example EN']),
      Example_KR: getCell(row, headerMap, ['Example_KR', 'Example KR']),
      Start_Time: getCell(row, headerMap, ['Start_Time', 'Start Time']),
      End_Time: getCell(row, headerMap, ['End_Time', 'End Time']),
      Drive_URL: audioInfo && audioInfo.Drive_URL ? audioInfo.Drive_URL : '',
      Audio_URL: audioInfo && audioInfo.Audio_URL ? audioInfo.Audio_URL : ''
    };

    words.push(item);
  }

  words.sort(function(a, b) {
    var aw = Number(a.Word_No || 0);
    var bw = Number(b.Word_No || 0);
    return aw - bw;
  });

  var wordsOnly = [];
  for (var w = 0; w < words.length; w++) {
    var copy = {};
    for (var k in words[w]) {
      if (k === 'Drive_URL' || k === 'Audio_URL') continue;
      copy[k] = words[w][k];
    }
    wordsOnly.push(copy);
  }
  wmCachePutJson_(cacheKey, wordsOnly, 21600);

  return words;
}


/* =========================================================
 * Word Mate Health Check
 * mode=health
 * 배포 /exec, HTML 파일, 운영 DB, 콘텐츠 DB, 샘플 세트/음원 상태를 한 번에 진단합니다.
 * Map.html / Study.html / 로그인 HTML은 수정하지 않습니다.
 * ========================================================= */
function healthCheck(e) {
  var startedAt = new Date();
  var sampleSetId = 'WM5-1-1';

  if (e && e.parameter && e.parameter.set_id) {
    sampleSetId = normalizeStudySetId(e.parameter.set_id);
  } else if (e && e.parameter && e.parameter.setId) {
    sampleSetId = normalizeStudySetId(e.parameter.setId);
  }

  var result = {
    success: true,
    mode: 'health',
    version: 'WM_HEALTH_CHECK_V1',
    checkedAt: Utilities.formatDate(startedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    execUrl: '',
    request: {
      sampleSetId: sampleSetId,
      studentId: e && e.parameter && e.parameter.studentId ? String(e.parameter.studentId || '').trim().toUpperCase() : ''
    },
    htmlFiles: {},
    lmsDb: {},
    contentDb: {},
    sampleSet: {},
    notes: []
  };

  try {
    result.execUrl = ScriptApp.getService().getUrl();
  } catch (errExec) {
    result.success = false;
    result.execUrl = '';
    result.notes.push('execUrl 확인 실패: ' + String(errExec && errExec.message ? errExec.message : errExec));
  }

  result.htmlFiles.Integrated = checkHtmlFileHealth_('Integrated');
  result.htmlFiles.Index = checkHtmlFileHealth_('Index');
  result.htmlFiles.Map = checkHtmlFileHealth_('Map.html');
  result.htmlFiles.Study = checkHtmlFileHealth_('Study');
  result.htmlFiles.Dev = checkHtmlFileHealth_('Dev.html');

  try {
    var ss = getLmsSpreadsheet_();
    if (!ss) {
      result.success = false;
      result.lmsDb.connected = false;
      result.lmsDb.message = 'SpreadsheetApp.openById(WM_LMS_SPREADSHEET_ID) 실패';
    } else {
      result.lmsDb.connected = true;
      result.lmsDb.spreadsheetName = ss.getName();
      result.lmsDb.sheets = {
        studentDb: checkSheetHealth_(ss, '1.학생관리_DB'),
        learningRecordDb: checkSheetHealth_(ss, '2.학습기록_DB'),
        commentDb: checkSheetHealth_(ss, '3.코멘트_DB'),
        reportDb: checkSheetHealth_(ss, '4.성적표생성_DB'),
        sendReadyDb: checkSheetHealth_(ss, '5.발송대기_DB'),
        sendLogDb: checkSheetHealth_(ss, '6.발송기록_DB'),
        classDb: checkSheetHealth_(ss, '7-1.반관리_DB')
      };
    }
  } catch (errLms) {
    result.success = false;
    result.lmsDb.connected = false;
    result.lmsDb.message = 'LMS DB 확인 실패: ' + String(errLms && errLms.message ? errLms.message : errLms);
  }

  try {
    var contentSS = SpreadsheetApp.openById(WM_CONTENT_SPREADSHEET_ID);
    result.contentDb.connected = true;
    result.contentDb.spreadsheetName = contentSS.getName();

    var parsed = parseWordMateSetId(sampleSetId);
    if (!parsed) {
      result.success = false;
      result.sampleSet.success = false;
      result.sampleSet.message = 'sampleSetId 형식 오류';
    } else {
      var levelSheetName = 'Level_' + pad2(parsed.level);
      var audioSheetName = 'L' + parsed.level;
      result.contentDb.levelSheet = checkSheetHealth_(contentSS, levelSheetName);
      result.contentDb.audioSheet = checkSheetHealth_(contentSS, audioSheetName);

      var levelSheet = contentSS.getSheetByName(levelSheetName);
      var audioInfo = getAudioInfoForSet(contentSS, audioSheetName, sampleSetId);
      var words = levelSheet ? getWordsForSet(levelSheet, sampleSetId, audioInfo) : [];
      var audioUrl = audioInfo && audioInfo.Audio_URL ? String(audioInfo.Audio_URL || '').trim() : '';

      result.sampleSet = {
        success: !!(words && words.length && audioUrl),
        setId: sampleSetId,
        wordCount: words ? words.length : 0,
        hasAudioUrl: !!audioUrl,
        audioUrlPreview: audioUrl ? audioUrl.substring(0, 80) + (audioUrl.length > 80 ? '...' : '') : '',
        firstWord: words && words.length ? String(words[0].Word || '') : '',
        message: words && words.length && audioUrl ? '샘플 세트/음원 조회 성공' : '샘플 세트 또는 Audio_URL 확인 필요'
      };

      if (!result.sampleSet.success) {
        result.success = false;
      }
    }
  } catch (errContent) {
    result.success = false;
    result.contentDb.connected = false;
    result.contentDb.message = '콘텐츠 DB 확인 실패: ' + String(errContent && errContent.message ? errContent.message : errContent);
  }

  result.summary = buildHealthSummary_(result);

  return outputResult(e, result);
}

function checkHtmlFileHealth_(fileName) {
  try {
    var content = HtmlService.createHtmlOutputFromFile(fileName).getContent();
    return {
      exists: true,
      fileName: fileName,
      length: content ? content.length : 0,
      message: 'OK'
    };
  } catch (err) {
    return {
      exists: false,
      fileName: fileName,
      length: 0,
      message: String(err && err.message ? err.message : err)
    };
  }
}

function checkSheetHealth_(ss, sheetName) {
  try {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return {
        exists: false,
        sheetName: sheetName,
        rows: 0,
        columns: 0,
        message: '시트 없음'
      };
    }

    return {
      exists: true,
      sheetName: sheetName,
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn(),
      message: 'OK'
    };
  } catch (err) {
    return {
      exists: false,
      sheetName: sheetName,
      rows: 0,
      columns: 0,
      message: String(err && err.message ? err.message : err)
    };
  }
}

function buildHealthSummary_(result) {
  var items = [];

  items.push(result.execUrl ? 'execUrl: OK' : 'execUrl: 확인 필요');
  items.push(result.htmlFiles && result.htmlFiles.Integrated && result.htmlFiles.Integrated.exists ? 'Integrated.html: OK' : 'Integrated.html: 확인 필요');
  items.push(result.htmlFiles && result.htmlFiles.Map && result.htmlFiles.Map.exists ? 'Map.html: OK' : 'Map.html: 확인 필요');
  items.push(result.htmlFiles && result.htmlFiles.Study && result.htmlFiles.Study.exists ? 'Study: OK' : 'Study: 확인 필요');
  items.push(result.lmsDb && result.lmsDb.connected ? 'LMS DB: OK' : 'LMS DB: 확인 필요');
  items.push(result.contentDb && result.contentDb.connected ? '콘텐츠 DB: OK' : '콘텐츠 DB: 확인 필요');
  items.push(result.sampleSet && result.sampleSet.success ? '샘플 세트/음원: OK' : '샘플 세트/음원: 확인 필요');

  return items.join(' / ');
}



/* WM_INTEGRATED_DIRECT_SEPARATED_ROUTE_V1
 * mode=integrated&screen=index|map|dev|study 요청을 Integrated.html router가 아니라
 * 현재 정상 확인 중인 분리 파일로 직접 연결합니다.
 */
function renderIntegratedDirectSeparatedApp_(e, mode) {
  var params = e && e.parameter ? e.parameter : {};
  var screen = String(params.screen || params.view || '').trim().toLowerCase();

  if (!screen) {
    if (mode === 'integratedMap') {
      screen = 'map';
    } else if (mode === 'integratedDev') {
      screen = 'dev';
    } else if (mode === 'integratedStudy') {
      screen = 'study';
    } else {
      screen = 'index';
    }
  }

  if (['index', 'map', 'dev', 'study'].indexOf(screen) === -1) {
    screen = 'index';
  }

  var nextParams = {};
  for (var key in params) {
    if (params.hasOwnProperty(key)) {
      nextParams[key] = params[key];
    }
  }

  if (screen === 'map') {
    nextParams.mode = 'map';
    return doGet({ parameter: nextParams });
  }

  if (screen === 'study') {
    nextParams.mode = 'study';
    return doGet({ parameter: nextParams });
  }

  if (screen === 'dev') {
    nextParams.mode = 'dev';
    return doGet({ parameter: nextParams });
  }

  var execUrl = '';
  try {
    execUrl = ScriptApp.getService().getUrl();
  } catch (execUrlErr) {
    execUrl = '';
  }

  var indexHtml = HtmlService.createHtmlOutputFromFile('Index').getContent();
  var indexBootScript = [
    '<script>',
    'window.WM_APPS_SCRIPT_URL = ' + JSON.stringify(execUrl) + ';',
    'window.WM_EXEC_URL = ' + JSON.stringify(execUrl) + ';',
    'try {',
    '  sessionStorage.setItem("WM_APPS_SCRIPT_URL", ' + JSON.stringify(execUrl) + ');',
    '  sessionStorage.setItem("WM_EXEC_URL", ' + JSON.stringify(execUrl) + ');',
    '} catch (e) {}',
    '</' + 'script>'
  ].join('\n');

  if (indexHtml.indexOf('</head>') !== -1) {
    indexHtml = indexHtml.replace('</head>', indexBootScript + '\n</head>');
  } else {
    indexHtml = indexBootScript + '\n' + indexHtml;
  }

  return HtmlService
    .createHtmlOutput(indexHtml)
    .setTitle('Word Mate Login');
}

/* =========================================================
 * Word Mate Integrated App Boot API
 * mode=integrated&screen=index|map|dev|study
 * 기존 Index/Map/Dev/Study는 그대로 두고, 새 Integrated.html만 별도 통합 테스트합니다.
 * ========================================================= */
function renderIntegratedApp(e, mode) {
  try {
    var params = e && e.parameter ? e.parameter : {};
    var execUrl = '';

    try {
      execUrl = ScriptApp.getService().getUrl();
    } catch (urlErr) {
      execUrl = '';
    }

    var screen = String(params.screen || params.view || '').trim().toLowerCase();

    if (!screen) {
      if (mode === 'integratedMap') {
        screen = 'map';
      } else if (mode === 'integratedDev') {
        screen = 'dev';
      } else if (mode === 'integratedStudy') {
        screen = 'study';
      } else {
        screen = 'index';
      }
    }

    if (['index', 'map', 'dev', 'study'].indexOf(screen) === -1) {
      screen = 'index';
    }

    var studentId = String(params.studentId || params.studentID || params.sid || params.wmStudentId || '').trim();
    var studentName = String(params.studentName || params.name || params.wmStudentName || params.wmName || '').trim();
    var learningAssign = String(params.learningAssign || params.level || params.learningLevel || params.wmLearningAssign || '').trim();
    var setId = String(params.set_id || params.setId || params.Set_ID || params.currentSetId || params.currentSet || params.wmCurrentSet || '').trim();

    var mapPayload = parseMapStudentPayload_(params.wmStudentPayload || params.studentPayload || '');
    if (mapPayload) {
      if (!studentId) {
        studentId = String(mapPayload['학생ID'] || mapPayload.studentId || '').trim();
      }
      if (!studentName) {
        studentName = String(mapPayload['학생이름'] || mapPayload.studentName || mapPayload.name || '').trim();
      }
      if (!learningAssign) {
        learningAssign = String(mapPayload['학습배정'] || mapPayload.learningAssign || mapPayload.level || '').trim();
      }
      if (!setId) {
        setId = String(mapPayload['현재세트'] || mapPayload.currentSet || '').trim();
      }
    }

    var studentInfo = getStudentBasicInfoForMap_(studentId);
    if (!studentName && studentInfo.studentName) {
      studentName = studentInfo.studentName;
    }
    if (!learningAssign && studentInfo.learningAssign) {
      learningAssign = studentInfo.learningAssign;
    }
    if (!learningAssign) {
      learningAssign = '4레벨';
    }
    if (!setId) {
      setId = buildInitialSetIdFromLearningAssign(learningAssign);
    }

    var normalizedStudySetId = normalizeStudySetId(setId);

    var boot = {
      app: 'Word Mate Integrated',
      version: 'WM_INTEGRATED_BOOT_V1',
      execUrl: execUrl,
      mode: 'integrated',
      screen: screen,
      student: {
        학생ID: studentId,
        학생이름: studentName,
        학습배정: learningAssign,
        현재세트: setId,
        studySetId: normalizedStudySetId
      },
      urls: {
        index: execUrl + '?mode=integrated&screen=index',
        map: execUrl + '?mode=integrated&screen=map',
        dev: execUrl + '?mode=integrated&screen=dev',
        study: execUrl + '?mode=integrated&screen=study'
      },
      api: {
        getLearningMap: execUrl + '?mode=getLearningMap',
        getSetData: execUrl + '?action=getSetData',
        saveLearningRecord: execUrl + '?action=saveLearningRecord',
        health: execUrl + '?mode=health'
      },
      params: copyPublicParams_(params)
    };

    var html = HtmlService.createHtmlOutputFromFile('Integrated').getContent();

    var bootScript = [
      '<script>',
      'window.WM_EXEC_URL = ' + JSON.stringify(execUrl) + ';',
      'window.WM_INTEGRATED_BOOT = ' + JSON.stringify(boot) + ';',
      'window.WM_MAP_SERVER_STUDENT = ' + JSON.stringify(boot.student) + ';',
      'window.WM_LOGGED_IN_STUDENT = ' + JSON.stringify(boot.student) + ';',
      'window.WM_CURRENT_STUDENT = ' + JSON.stringify(boot.student) + ';',
      'window.WM_STUDENT_ID = ' + JSON.stringify(studentId) + ';',
      'window.WM_STUDENT_NAME = ' + JSON.stringify(studentName) + ';',
      'window.WM_LEARNING_ASSIGN = ' + JSON.stringify(learningAssign) + ';',
      'window.WM_INITIAL_SET_ID = ' + JSON.stringify(setId) + ';',
      'window.WM_INITIAL_STUDY_SET_ID = ' + JSON.stringify(normalizedStudySetId) + ';',
      'try {',
      '  sessionStorage.setItem("wmLoggedInStudent", ' + JSON.stringify(JSON.stringify(boot.student)) + ');',
      '  sessionStorage.setItem("WM_STUDENT_ID", ' + JSON.stringify(studentId) + ');',
      '  sessionStorage.setItem("WM_STUDENT_NAME", ' + JSON.stringify(studentName) + ');',
      '  sessionStorage.setItem("WM_LEARNING_ASSIGN", ' + JSON.stringify(learningAssign) + ');',
      '  sessionStorage.setItem("WM_INITIAL_SET_ID", ' + JSON.stringify(setId) + ');',
      '  sessionStorage.setItem("WM_EXEC_URL", ' + JSON.stringify(execUrl) + ');',
      '} catch (e) {}',
      '</' + 'script>'
    ].join('\n');

    html = html
      .replace(/__WM_EXEC_URL__/g, execUrl)
      .replace(/__WM_BOOT_JSON__/g, JSON.stringify(boot))
      .replace(/__WM_SCREEN__/g, screen)
      .replace(/__WM_STUDENT_ID__/g, studentId)
      .replace(/__WM_STUDENT_NAME__/g, studentName)
      .replace(/__WM_LEARNING_ASSIGN__/g, learningAssign)
      .replace(/__WM_INITIAL_SET_ID__/g, setId)
      .replace(/__WM_INITIAL_STUDY_SET_ID__/g, normalizedStudySetId);

    /*
     * WM_INTEGRATED_BOOT_INJECTION_FIX_V2
     * Integrated.html은 실제 top-level <head>가 없고,
     * textarea 내부 Index 템플릿에 </head>가 먼저 등장합니다.
     * 따라서 html.replace('</head>', ...)를 사용하면 bootScript가
     * 실행 문서가 아니라 textarea 템플릿 내부로 들어갈 수 있습니다.
     * bootScript는 router 실행 전에 top-level script로 주입합니다.
     */
    var routerMarkerV16 = '<script id="wm-integrated-router-v16-no-document-write">';
    var routerMarkerV12 = '<script id="wm-integrated-router-v12-domparser">';
    var routerMarkerV11 = '<script id="wm-integrated-router-v11">';
    var routerMarkerV7 = '<script id="wm-integrated-router-v7-early-polling">';
    var routerMarkerV5 = '<script id="wm-integrated-router-v5">';
    if (html.indexOf(routerMarkerV16) !== -1) {
      html = html.replace(routerMarkerV16, bootScript + '\n' + routerMarkerV16);
    } else if (html.indexOf(routerMarkerV12) !== -1) {
      html = html.replace(routerMarkerV12, bootScript + '\n' + routerMarkerV12);
    } else if (html.indexOf(routerMarkerV11) !== -1) {
      html = html.replace(routerMarkerV11, bootScript + '\n' + routerMarkerV11);
    } else if (html.indexOf(routerMarkerV7) !== -1) {
      html = html.replace(routerMarkerV7, bootScript + '\n' + routerMarkerV7);
    } else if (html.indexOf(routerMarkerV5) !== -1) {
      html = html.replace(routerMarkerV5, bootScript + '\n' + routerMarkerV5);
    } else if (html.indexOf('<div id="wm-integrated-router-loading"') !== -1) {
      html = html.replace('<div id="wm-integrated-router-loading"', bootScript + '\n<div id="wm-integrated-router-loading"');
    } else {
      html = bootScript + '\n' + html;
    }

    return HtmlService
      .createHtmlOutput(html)
      .setTitle('Word Mate Integrated');

  } catch (err) {
    return outputResult(e, {
      success: false,
      message: 'Integrated 화면 생성 오류',
      error: String(err && err.message ? err.message : err)
    });
  }
}

function copyPublicParams_(params) {
  var copied = {};
  params = params || {};

  for (var key in params) {
    if (!params.hasOwnProperty(key)) {
      continue;
    }
    if (String(key).toLowerCase() === 'password') {
      continue;
    }
    copied[key] = String(params[key] || '');
  }

  return copied;
}


/* WM_EXISTING_STUDENT_CLASS_DB_ONE_TIME_REPAIR_V1
 * 기존 학생관리_DB에서 반명이 저장된 학생 중
 * 7-2.반관리_DB 해당 반 명단에 누락된 학생만 1회 보정합니다.
 * 기존 반/교사/상태/등록일/등록자/메모 및 기존 학생명단은 삭제·덮어쓰기하지 않습니다.
 */
function wmRepairMissingStudentsToClassDbOnce_() {
  var ss = getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  var classSheet = ss.getSheetByName('7-2.반관리_DB');

  if (!studentSheet) throw new Error('1.학생관리_DB 시트를 찾을 수 없습니다.');
  if (!classSheet) throw new Error('7-2.반관리_DB 시트를 찾을 수 없습니다.');
  if (studentSheet.getLastRow() < 2 || classSheet.getLastRow() < 2) {
    return {success:true, updatedClasses:0, addedStudents:0, skippedStudents:0, message:'보정할 데이터가 없습니다.'};
  }

  function norm_(value) {
    return String(value || '').trim().toUpperCase();
  }
  function splitList_(value) {
    return String(value || '')
      .split(/[,\n\r;|]+/)
      .map(function(v){ return String(v || '').trim(); })
      .filter(Boolean);
  }

  var studentValues = studentSheet.getDataRange().getDisplayValues();
  var studentHeaders = studentValues[0].map(function(v){ return String(v || '').trim(); });
  var idxStudentId = wmFindHeaderIndex_(studentHeaders, ['학생ID','Student_ID','studentId','student_id','StudentID']);
  var idxStudentName = wmFindHeaderIndex_(studentHeaders, ['학생이름','학생명','이름','studentName','Student_Name','StudentName']);
  var idxStudentClass = wmFindHeaderIndex_(studentHeaders, ['반명','Class','반','className','class']);

  if (idxStudentId < 0 || idxStudentName < 0 || idxStudentClass < 0) {
    throw new Error('1.학생관리_DB 필수 컬럼(학생ID/학생이름/반명)을 찾을 수 없습니다.');
  }

  var classValues = classSheet.getDataRange().getDisplayValues();
  var classHeaders = classValues[0].map(function(v){ return String(v || '').trim(); });
  var idxClassName = classHeaders.indexOf('반명');
  var idxClassStudentIds = classHeaders.indexOf('학생ID목록');
  var idxClassStudentNames = classHeaders.indexOf('학생이름목록');
  var idxClassStudentCount = classHeaders.indexOf('학생수');
  var idxClassUpdatedAt = classHeaders.indexOf('수정일');

  if (idxClassName < 0 || idxClassStudentIds < 0 || idxClassStudentNames < 0 || idxClassStudentCount < 0) {
    throw new Error('7-2.반관리_DB 필수 컬럼(반명/학생ID목록/학생이름목록/학생수)을 찾을 수 없습니다.');
  }

  /* 같은 반명이 중복된 경우 현재 화면/신규등록 동작과 동일하게 마지막 행을 사용합니다. */
  var classRowByKey = {};
  for (var cr = 1; cr < classValues.length; cr++) {
    var classKey = norm_(classValues[cr][idxClassName]);
    if (classKey) classRowByKey[classKey] = cr;
  }

  var pendingByClass = {};
  var skippedStudents = 0;
  for (var sr = 1; sr < studentValues.length; sr++) {
    var studentId = String(studentValues[sr][idxStudentId] || '').trim();
    var studentName = String(studentValues[sr][idxStudentName] || '').trim();
    var className = String(studentValues[sr][idxStudentClass] || '').trim();
    var studentKey = norm_(studentId);
    var classKey = norm_(className);

    if (!studentKey || !classKey || !classRowByKey.hasOwnProperty(classKey)) {
      skippedStudents++;
      continue;
    }
    if (!pendingByClass[classKey]) pendingByClass[classKey] = [];
    pendingByClass[classKey].push({id:studentId, name:studentName});
  }

  var updatedClasses = 0;
  var addedStudents = 0;
  Object.keys(pendingByClass).forEach(function(classKey){
    var rowIndex = classRowByKey[classKey];
    var existingIds = splitList_(classValues[rowIndex][idxClassStudentIds]);
    var existingNames = splitList_(classValues[rowIndex][idxClassStudentNames]);
    var idSet = {};
    existingIds.forEach(function(id){ idSet[norm_(id)] = true; });

    var classAdded = 0;
    pendingByClass[classKey].forEach(function(student){
      var key = norm_(student.id);
      if (!key || idSet[key]) return;
      existingIds.push(student.id);
      existingNames.push(student.name || '');
      idSet[key] = true;
      classAdded++;
      addedStudents++;
    });

    if (!classAdded) return;

    var sheetRow = rowIndex + 1;
    classSheet.getRange(sheetRow, idxClassStudentIds + 1).setValue(existingIds.join(', '));
    classSheet.getRange(sheetRow, idxClassStudentNames + 1).setValue(existingNames.join(', '));
    classSheet.getRange(sheetRow, idxClassStudentCount + 1).setValue(existingIds.length);
    if (idxClassUpdatedAt >= 0) classSheet.getRange(sheetRow, idxClassUpdatedAt + 1).setValue(new Date());
    updatedClasses++;
  });

  SpreadsheetApp.flush();
  return {
    success:true,
    updatedClasses:updatedClasses,
    addedStudents:addedStudents,
    skippedStudents:skippedStudents,
    message:'기존 누락 학생 반관리_DB 1회 보정 완료'
  };
}

/* =========================================================
 * WM_GRADE_DB_PRIMARY_DATA_AUTO_SYNC_20260729_V1
 * 성적표 예약 발행 또는 명시적 복구 시에만 실행하며, 실제 발행행이 있는 기간만 계산합니다.
 * 대상: 학생관리_DB에 출석약속·세트약속이 모두 있고,
 *       2.학습기록_DB에 실제 학습기록이 있는 학생.
 * 출력: 6.성적등급_DB A:AH를 갱신합니다.
 *       J=학습시작일, K=학습종료일, U=출석률, V=세트진행률, W=출석점수
 *       X=1세트평균시간, Y=참여집중도, Z=원점수환산점수
 *       AA=표준점수, AB=점수추이, AC=최종점수, AD=Grade, AE=Class
 *       AF=자동코멘트, AG=성적표반영상태, AH=생성일시
 *       AG가 반영완료인 기존 행은 AF/AG/AH를 그대로 보존합니다.
 * 집계키: 학생ID + 연도 + 월 + 주차
 * ========================================================= */
function wmSyncGradePrimaryDataAfterLearningSave_(ss) {
  ss = ss || getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  var recordSheet = ss.getSheetByName('2.학습기록_DB');
  var gradeSheet = ss.getSheetByName('6.성적등급_DB');

  if (!studentSheet || !recordSheet || !gradeSheet) {
    return {success:false, message:'학생관리_DB / 학습기록_DB / 6.성적등급_DB 시트를 확인하세요.'};
  }

  var targetHeaders = [
    '일련번호','학생ID','학생이름','반명','교사명','연도','월','주차','리포트유형',
    '학습시작일','학습종료일','출석약속','세트약속','실제출석일수','완료세트수','학습단어수','총학습시간',
    '평균원점수','등록일','수정일','출석률','세트진행률','출석점수',
    '1세트평균시간','참여집중도','원점수환산점수','표준점수','점수추이',
    '최종점수','Grade','Class','자동코멘트','성적표반영상태','생성일시'
  ];
  var actualTargetHeaders = gradeSheet.getRange(1, 1, 1, targetHeaders.length).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  for (var th = 0; th < targetHeaders.length; th++) {
    if (actualTargetHeaders[th] !== targetHeaders[th]) {
      return {success:false, message:'6.성적등급_DB A:AH 헤더 불일치: ' + (th + 1) + '열 ' + actualTargetHeaders[th]};
    }
  }

  function norm_(value) {
    return String(value == null ? '' : value).trim();
  }
  function sid_(value) {
    return norm_(value).toUpperCase();
  }
  function idx_(headers, names) {
    for (var i = 0; i < names.length; i++) {
      var found = headers.indexOf(names[i]);
      if (found >= 0) return found;
    }
    return -1;
  }
  function cell_(row, index) {
    return index >= 0 ? row[index] : '';
  }
  function number_(value) {
    var text = norm_(value).replace(/,/g, '').replace(/[^0-9.\-]/g, '');
    var number = Number(text);
    return text !== '' && isFinite(number) ? number : 0;
  }
  function round1_(value) {
    return Math.round(Number(value || 0) * 10) / 10;
  }
  function dateParts_(value) {
    var text = norm_(value);
    var match = text.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (!match) return null;
    var y = Number(match[1]);
    var m = Number(match[2]);
    var d = Number(match[3]);
    if (!y || !m || !d) return null;
    var iso = y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
    var week = typeof wmReportCalendarWeek_ === 'function' ? wmReportCalendarWeek_(iso) : Math.ceil(d / 7);
    week = Math.max(1, Math.min(5, Number(week || 1)));
    return {year:y, month:m, day:d, iso:iso, week:week};
  }
  function seconds_(value) {
    if (typeof wmReportParseSeconds_ === 'function') return Number(wmReportParseSeconds_(value) || 0);
    var text = norm_(value);
    if (!text) return 0;
    if (/^\d+(?:\.\d+)?$/.test(text)) return Number(text);
    var total = 0, match;
    match = text.match(/(\d+)\s*시간/); if (match) total += Number(match[1]) * 3600;
    match = text.match(/(\d+)\s*분/); if (match) total += Number(match[1]) * 60;
    match = text.match(/(\d+)\s*초/); if (match) total += Number(match[1]);
    if (text.indexOf(':') >= 0) {
      var parts = text.split(':').map(Number);
      if (parts.length === 3) total = parts[0] * 3600 + parts[1] * 60 + parts[2];
      if (parts.length === 2) total = parts[0] * 60 + parts[1];
    }
    return isFinite(total) ? total : 0;
  }
  function timeText_(seconds) {
    seconds = Math.max(0, Math.round(Number(seconds || 0)));
    var h = Math.floor(seconds / 3600);
    var m = Math.floor((seconds % 3600) / 60);
    var s = seconds % 60;
    if (h) return h + '시간 ' + m + '분 ' + s + '초';
    if (m) return m + '분 ' + s + '초';
    return s + '초';
  }
  function completed_(row, idxStatus, idxComplete) {
    var status = norm_(cell_(row, idxStatus)).toUpperCase();
    var complete = norm_(cell_(row, idxComplete)).toUpperCase();
    return status === '완료' || status === 'COMPLETE' || status === 'COMPLETED' ||
      complete === '완료' || complete === 'TRUE' || complete === 'Y' || complete === '1';
  }
  function wordCount_(setId) {
    if (typeof wmReportWordCountForSet_ === 'function') return Number(wmReportWordCountForSet_(setId) || 0);
    var match = norm_(setId).match(/WM(\d+)-/i);
    var level = match ? Number(match[1]) : 0;
    return level >= 3 && level <= 6 ? 10 : (level >= 7 && level <= 13 ? 20 : 0);
  }

  function isMonthlyReport_(reportType) {
    var type = norm_(reportType).replace(/\s+/g, '');
    return type === '월간' || type === 'MONTHLY' || type === 'Monthly';
  }
  function gradeGoalMessage_(grade, reportType) {
    var nextPeriod = isMonthlyReport_(reportType) ? '다음 달' : '다음 주';
    if (Number(grade) === 1) {
      return nextPeriod + '에도 현재 학습 리듬을 유지하며 최고 등급에 도전해 보세요.';
    }
    if (Number(grade) === 2) {
      return nextPeriod + '에는 출석과 세트학습을 꾸준히 실천하여 1등급에 도전해 보세요.';
    }
    if (Number(grade) === 3) {
      return nextPeriod + '에는 약속한 학습계획을 꾸준히 실천하여 더 높은 등급에 도전해 보세요.';
    }
    return '';
  }
  function buildAutoComment_(data) {
    return wmBuildGradeAutoComment_(data);
  }

  var studentValues = studentSheet.getDataRange().getDisplayValues();
  var recordValues = recordSheet.getDataRange().getDisplayValues();
  if (studentValues.length < 2 || recordValues.length < 2) {
    /* 원천 데이터가 일시적으로 비어도 이미 발행된 성적등급_DB는 절대 삭제하지 않습니다. */
    return {success:true, rows:Math.max(0, gradeSheet.getLastRow() - 1), preserved:true, message:'대상 데이터 없음 · 기존 성적등급_DB 보존'};
  }

  var sh = studentValues[0].map(function(v){ return norm_(v); });
  var rh = recordValues[0].map(function(v){ return norm_(v); });
  var sId = idx_(sh, ['학생ID','Student_ID','studentId']);
  var sName = idx_(sh, ['학생이름','학생명','이름']);
  var sClass = idx_(sh, ['반명','Class','반','학급명']);
  var sTeacher = idx_(sh, ['교사명','담당교사','Teacher']);
  var sAttendancePromise = idx_(sh, ['출석약속','주간출석약속','출석일수약속','약속출석일수']);
  var sSetPromise = idx_(sh, ['세트약속','주간세트약속','완료세트약속','약속세트수']);
  var sLearningStart = idx_(sh, ['학습시작일']);
  var sLearningEnd = idx_(sh, ['학습종료일']);
  var sRegistrationDate = idx_(sh, ['등록일']);

  var rId = idx_(rh, ['학생ID','Student_ID','studentId']);
  var rRecordId = idx_(rh, ['학습기록ID','Learning_Record_ID','recordId']);
  var rName = idx_(rh, ['학생이름','학생명','이름']);
  var rClass = idx_(rh, ['Class','반명','반','학급명']);
  var rTeacher = idx_(rh, ['교사명','담당교사','Teacher']);
  var rDate = idx_(rh, ['학습날짜','학습일','날짜','등록일']);
  var rSet = idx_(rh, ['Set_ID','Set ID','세트ID']);
  var rScore = idx_(rh, ['점수','한영주관식','원점수']);
  var rTime = idx_(rh, ['총소요시간','총학습시간','총 소요시간']);
  var rStatus = idx_(rh, ['완료상태','상태']);
  var rComplete = idx_(rh, ['완료']);

  if (sId < 0 || sAttendancePromise < 0 || sSetPromise < 0 || rId < 0 || rDate < 0) {
    return {success:false, message:'필수 컬럼 누락: 학생ID/출석약속/세트약속/학습날짜'};
  }

  var students = {};
  for (var s = 1; s < studentValues.length; s++) {
    var srow = studentValues[s];
    var studentId = sid_(cell_(srow, sId));
    var attendancePromise = norm_(cell_(srow, sAttendancePromise));
    var setPromise = norm_(cell_(srow, sSetPromise));
    if (!studentId || !attendancePromise || !setPromise) continue;
    students[studentId] = {
      studentId:studentId,
      studentName:norm_(cell_(srow, sName)),
      className:norm_(cell_(srow, sClass)),
      teacherName:norm_(cell_(srow, sTeacher)),
      attendancePromise:attendancePromise,
      setPromise:setPromise,
      learningStartDate:norm_(cell_(srow, sLearningStart)) || norm_(cell_(srow, sRegistrationDate)),
      learningEndDate:norm_(cell_(srow, sLearningEnd))
    };
  }

  var groups = {};
  var completedRows = [];

  function addCompletedRecordToGroup_(groupKey, reportType, weekNumber, dp, row, recordIndex, student) {
    if (!groups[groupKey]) {
      groups[groupKey] = {
        key:groupKey, studentId:sid_(cell_(row, rId)),
        studentName:student.studentName || norm_(cell_(row, rName)),
        className:student.className || norm_(cell_(row, rClass)),
        teacherName:student.teacherName || norm_(cell_(row, rTeacher)),
        year:dp.year, month:dp.month, week:weekNumber, reportType:reportType,
        attendancePromise:student.attendancePromise, setPromise:student.setPromise,
        learningStartDate:student.learningStartDate, learningEndDate:student.learningEndDate,
        dates:{}, completedSetIds:{}, completedSets:0, wordCount:0,
        totalSeconds:0, scoreSum:0, scoreCount:0,
        weeks:[
          {dates:{},setIds:{},attendanceDays:0,completedSets:0},
          {dates:{},setIds:{},attendanceDays:0,completedSets:0},
          {dates:{},setIds:{},attendanceDays:0,completedSets:0},
          {dates:{},setIds:{},attendanceDays:0,completedSets:0},
          {dates:{},setIds:{},attendanceDays:0,completedSets:0}
        ]
      };
    }
    var group = groups[groupKey];
    group.dates[dp.iso] = true;
    var groupWeek = group.weeks[Math.max(1, Math.min(5, Number(dp.week || 1))) - 1];
    groupWeek.dates[dp.iso] = true;
    groupWeek.attendanceDays = Object.keys(groupWeek.dates).length;
    group.totalSeconds += seconds_(cell_(row, rTime));
    var scoreText = norm_(cell_(row, rScore));
    var score = Number(String(scoreText).replace(/[^0-9.\-]/g, ''));
    if (scoreText !== '' && isFinite(score)) {
      group.scoreSum += score;
      group.scoreCount += 1;
    }
    var setId = norm_(cell_(row, rSet));
    var learningRecordId = norm_(cell_(row, rRecordId));
    var completedKey = learningRecordId || ('ROW_' + recordIndex);
    if (!group.completedSetIds[completedKey]) {
      group.completedSetIds[completedKey] = true;
      group.completedSets += 1;
      group.wordCount += wordCount_(setId);
    }
    if (!groupWeek.setIds[completedKey]) {
      groupWeek.setIds[completedKey] = true;
      groupWeek.completedSets += 1;
    }
  }

  for (var r = 1; r < recordValues.length; r++) {
    var row = recordValues[r];
    var recordStudentId = sid_(cell_(row, rId));
    var student = students[recordStudentId];
    if (!student || !completed_(row, rStatus, rComplete)) continue;
    var dp = dateParts_(cell_(row, rDate));
    if (!dp) continue;
    var studentLearningStart = wmReportNormalizeDate_(student.learningStartDate);
    var studentLearningEnd = wmReportNormalizeDate_(student.learningEndDate);
    if (studentLearningStart && dp.iso < studentLearningStart) continue;
    if (studentLearningEnd && dp.iso > studentLearningEnd) continue;
    var periodInfo = wmReportPeriodInfo_(dp.iso);
    if (!periodInfo) continue;
    completedRows.push({row:row, index:r, student:student, dp:dp, periodInfo:periodInfo});
  }

  completedRows.forEach(function(item) {
    var dp = item.dp;
    var recordStudentId = sid_(cell_(item.row, rId));
    var yearMonth = dp.year + '-' + ('0' + dp.month).slice(-2);
    var monthEndInfo = wmReportPeriodInfo_(yearMonth + '-' + ('0' + new Date(dp.year, dp.month, 0).getDate()).slice(-2));

    /* Weekly는 별도 Weekly 결과를 더하지 않고 2.학습기록_DB 누적기록을 한 번 집계합니다.
     * 해당 월 1일부터 선택한 종료 주차 말일까지의 완료기록이 각 Weekly 한 행을 구성합니다. */
    var lastWeek = monthEndInfo ? Number(monthEndInfo.lastWeek || 0) : 0;
    for (var targetWeek = Number(dp.week || 1); targetWeek < lastWeek; targetWeek++) {
      var targetRange = wmReportCalendarWeekRange_(yearMonth, targetWeek);
      var targetPeriodInfo = wmReportPeriodInfo_(targetRange.endDate);
      if (!targetPeriodInfo || targetPeriodInfo.isLastWeek || !wmReportIsReleased_(targetPeriodInfo)) continue;
      var weeklyKey = recordStudentId + '|' + dp.year + '|' + dp.month + '|' + targetWeek;
      addCompletedRecordToGroup_(weeklyKey, '주간', targetWeek, dp, item.row, item.index, item.student);
    }

    /* 월이 종료되면 1일~말일 전체 완료 세트를 마지막 주차의 Monthly 한 행으로 합산합니다. */
    var monthlyLearningStart = wmReportNormalizeDate_(item.student.learningStartDate);
    var monthlyLearningEnd = wmReportNormalizeDate_(item.student.learningEndDate);
    if (monthEndInfo && wmReportIsReleased_(monthEndInfo) && (!monthlyLearningStart || dp.iso >= monthlyLearningStart) && (!monthlyLearningEnd || dp.iso <= monthlyLearningEnd)) {
      var monthlyKey = recordStudentId + '|' + dp.year + '|' + dp.month + '|' + monthEndInfo.lastWeek;
      addCompletedRecordToGroup_(monthlyKey, '월간', monthEndInfo.lastWeek, dp, item.row, item.index, item.student);
    }
  });

  var reportedGradeKeys = {};
  var reportSheetForStatus = ss.getSheetByName('4-1.성적표생성_DB');
  if (reportSheetForStatus && reportSheetForStatus.getLastRow() >= 2) {
    var reportStatusValues = reportSheetForStatus.getDataRange().getDisplayValues();
    var reportStatusHeaders = reportStatusValues[0].map(function(v){ return norm_(v); });
    var rsStudent = idx_(reportStatusHeaders, ['학생ID','Student_ID','studentId']);
    var rsWeek = idx_(reportStatusHeaders, ['리포트주차']);
    for (var rr = 1; rr < reportStatusValues.length; rr++) {
      var rsSid = sid_(cell_(reportStatusValues[rr], rsStudent));
      var rsWeekText = norm_(cell_(reportStatusValues[rr], rsWeek));
      var rsMatch = rsWeekText.match(/^(\d{4})-(\d{1,2})-(\d+)주차$/);
      if (rsSid && rsMatch) {
        reportedGradeKeys[rsSid + '|' + Number(rsMatch[1]) + '|' + Number(rsMatch[2]) + '|' + Number(rsMatch[3])] = true;
      }
    }
  }

  var existingRegistered = {};
  var existingPublishedRows = {};
  var existingEditableRows = {};
  var editableRowPool = [];
  var existingLastRow = gradeSheet.getLastRow();
  if (existingLastRow >= 2) {
    var existingRange = gradeSheet.getRange(2, 1, existingLastRow - 1, targetHeaders.length);
    var existingRaw = existingRange.getValues();
    var existing = existingRange.getDisplayValues();
    for (var er = 0; er < existing.length; er++) {
      var existingKey = sid_(existing[er][1]) + '|' + Number(existing[er][5] || 0) + '|' + Number(existing[er][6] || 0) + '|' + Number(existing[er][7] || 0);
      if (existingKey !== '|0|0|0') {
        if (existing[er][18]) existingRegistered[existingKey] = existing[er][18];
        var existingReportStatus = norm_(existing[er][32]);
        var existingRowNumber = er + 2;
        if (existingReportStatus === '반영완료') {
          existingPublishedRows[existingKey] = {rowNumber:existingRowNumber, row:existingRaw[er].slice(), reportType:norm_(existing[er][8])};
        } else {
          existingEditableRows[existingKey] = {rowNumber:existingRowNumber, row:existingRaw[er].slice(), reportType:norm_(existing[er][8])};
          editableRowPool.push(existingRowNumber);
        }
      }
    }
  }

  var now = new Date();
  var rows = Object.keys(groups).map(function(key){ return groups[key]; });
  rows.sort(function(a,b){
    return b.year - a.year || b.month - a.month || b.week - a.week || a.studentId.localeCompare(b.studentId);
  });
  var pendingRows = [];
  rows.forEach(function(group){
    var reportType = group.reportType || '주간';
    /* 성적등급 계산은 4-1.성적표생성_DB에 실제 발행행이 만들어진 기간에만 실행합니다. */
    if (!reportedGradeKeys[group.key]) return;
    /* 반영완료 행은 읽기만 하며 어떤 동기화에서도 삭제·재계산·덮어쓰기하지 않습니다. */
    if (existingPublishedRows[group.key]) {
      if (existingPublishedRows[group.key].reportType !== String(reportType)) {
        throw new Error('발행완료 성적표 유형 충돌: ' + group.key + ' / ' + existingPublishedRows[group.key].reportType + ' / ' + reportType);
      }
      return;
    }

    var actualAttendanceDays = Object.keys(group.dates).length;
    var attendancePromiseNumber = number_(group.attendancePromise);
    var setPromiseNumber = number_(group.setPromise);
    /* 주간·월간 모두 성적표보기와 같은 단일 월 누적 계산함수를 사용합니다. */
    var calculatedGrade = wmCalculateMonthlyCumulativeGrade_({
      yearMonth:group.year + '-' + ('0' + group.month).slice(-2),
      selectedWeek:group.week,
      learningStartDate:group.learningStartDate,
      learningEndDate:group.learningEndDate,
      attendancePromise:attendancePromiseNumber,
      setPromise:setPromiseNumber,
      weeks:group.weeks,
      completedSets:group.completedSets,
      totalSeconds:group.totalSeconds,
      scoreSum:group.scoreSum,
      scoreCount:group.scoreCount
    });
    var attendanceRate = calculatedGrade.attendanceRate;
    var setProgressRate = calculatedGrade.setProgressRate;
    var attendanceScore = calculatedGrade.attendanceScore;
    var averageSetSeconds = calculatedGrade.averageSetSeconds;
    var averageSetTime = group.completedSets > 0 ? timeText_(averageSetSeconds) : '';
    var participationScore = calculatedGrade.participationScore;
    var averageRawScore = calculatedGrade.averageRawScore;
    var convertedRawScore = calculatedGrade.convertedRawScore;
    var standardScore = calculatedGrade.standardScore;
    var scoreTrend = calculatedGrade.scoreTrend;
    var finalScore = calculatedGrade.finalScore;
    var grade = calculatedGrade.grade;
    var gradeClass = calculatedGrade.gradeClass;

    /* 신규 발행 행의 좌측 점수와 우측 자동코멘트는 같은 calculatedGrade 값으로 동시에 생성합니다. */
    var autoComment = buildAutoComment_({
          studentName:group.studentName,
          year:group.year,
          month:group.month,
          week:group.week,
          reportType:reportType,
          attendancePromise:group.attendancePromise,
          setPromise:group.setPromise,
          attendanceScore:attendanceScore,
          participationScore:participationScore,
          scoreTrend:scoreTrend,
          finalScore:finalScore,
          grade:grade,
          gradeClass:gradeClass
        });
    var reportStatus = '반영완료';
    var commentCreatedAt = now;

    pendingRows.push({key:group.key, row:[
      0, group.studentId, group.studentName, group.className, group.teacherName,
      group.year, group.month, group.week, reportType, group.learningStartDate, group.learningEndDate, group.attendancePromise, group.setPromise,
      actualAttendanceDays, group.completedSets, group.wordCount,
      timeText_(group.totalSeconds), averageRawScore,
      existingRegistered[group.key] || now, now,
      attendanceRate, setProgressRate, attendanceScore, averageSetTime, participationScore,
      convertedRawScore, standardScore, scoreTrend, finalScore, grade, gradeClass,
      autoComment, reportStatus, commentCreatedAt
    ]});
  });

  /* 미발행 행만 같은 키의 기존 행을 갱신하거나 빈 미발행 행을 재사용합니다. */
  var usedEditableRows = {};
  var nextAppendRow = Math.max(2, gradeSheet.getLastRow() + 1);
  pendingRows.forEach(function(item) {
    var rowNumber = existingEditableRows[item.key] ? existingEditableRows[item.key].rowNumber : 0;
    if (!rowNumber) {
      while (editableRowPool.length && usedEditableRows[editableRowPool[0]]) editableRowPool.shift();
      if (editableRowPool.length) rowNumber = editableRowPool.shift();
    }
    if (!rowNumber) rowNumber = nextAppendRow++;
    usedEditableRows[rowNumber] = true;
    item.row[0] = rowNumber - 1;
    gradeSheet.getRange(rowNumber, 1, 1, targetHeaders.length).setValues([item.row]);
  });

  /* 더 이상 필요한 미발행 행만 개별 삭제합니다. 반영완료 행은 절대 건드리지 않습니다. */
  Object.keys(existingEditableRows).forEach(function(key) {
    var rowNumber = existingEditableRows[key].rowNumber;
    if (!usedEditableRows[rowNumber]) {
      gradeSheet.getRange(rowNumber, 1, 1, targetHeaders.length).clearContent();
    }
  });

  SpreadsheetApp.flush();
  return {
    success:true,
    rows:Math.max(0, gradeSheet.getLastRow() - 1),
    updatedPendingRows:pendingRows.length,
    preservedPublishedRows:Object.keys(existingPublishedRows).length,
    message:'6.성적등급_DB 발행완료 행 보존 · 미발행 행만 갱신 완료'
  };
}

/* WM_LMS_SEND_CENTER_DB_CONNECTION_V1
 * 5-1.발송내역_DB와 5-2.자동발송설정_DB를 계산 없이 그대로 읽어 LMS에 반환합니다.
 * TEACHER/SUB_LEADER는 기존 학생 조회범위와 같은 학생·반의 행만 반환합니다.
 */
/* WM_SEND_CENTER_PRIMARY_DATA_DIRECT_WRITE_V5
 * 기존 DB와 같은 방식으로 학생관리_DB 저장·수정 실행 시 5-1·5-2 구글시트에 직접 기록합니다.
 * 5-1 학생정보·수신자정보는 1.학생관리_DB, 최신 리포트정보는 6.성적등급_DB에서 기록합니다.
 * 시트를 열 때는 기존 학생 전체를 직접 기록하고, 이후에는 학생·성적등급 저장 시 즉시 갱신합니다.
 */
function wmSyncSendCenterPrimaryDataFromStudents_(students, ss) {
  ss = ss || getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  var logSheet = ss.getSheetByName('5-1.발송내역_DB');
  var settingSheet = ss.getSheetByName('5-2.자동발송설정_DB');

  if (!logSheet) {
    logSheet = ss.getSheetByName(['5-1.발송내역','DB'].join(''));
    if (logSheet) logSheet.setName('5-1.발송내역_DB');
  }
  if (!settingSheet) {
    settingSheet = ss.getSheetByName(['5-2.자동발송설정','DB'].join(''));
    if (settingSheet) settingSheet.setName('5-2.자동발송설정_DB');
  }

  if (!studentSheet) throw new Error('1.학생관리_DB 시트를 찾을 수 없습니다.');
  if (!logSheet) throw new Error('5-1.발송내역_DB 시트를 찾을 수 없습니다.');
  if (!settingSheet) throw new Error('5-2.자동발송설정_DB 시트를 찾을 수 없습니다.');

  /* WM_SEND_CENTER_FINAL_HEADERS_20260808_V1
   * 5-1은 발송 결과 전용 14개 컬럼이며 학생ID 기반 사전행 자동생성을 하지 않습니다.
   * 5-2는 학생ID·학생이름을 학생관리_DB에서 직접 동기화합니다. */
  var logHeaders = ['교사명','리포트유형','리포트월','리포트주차','수신자구분','학부모명','학부모연락처','발송채널','실제발송일시','발송상태','성적표링크','실패사유','API메시지ID','비고'];
  var settingHeaders = ['학생ID','학생이름','리포트유형','발송주기','발송요일','발송일','발송시간','자동발송상태','비고','등록일시','수정일시'];
  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  if (logSheet.getLastColumn() > logHeaders.length) {
    logSheet.getRange(1, logHeaders.length + 1, 1, logSheet.getLastColumn() - logHeaders.length).clearContent();
  }
  settingSheet.getRange(1, 1, 1, settingHeaders.length).setValues([settingHeaders]);

  if (!Array.isArray(students)) {
    var studentValues = studentSheet.getDataRange().getDisplayValues();
    var studentHeaders = studentValues.length ? studentValues[0].map(function(value) { return String(value || '').trim(); }) : [];
    students = studentValues.slice(1).map(function(values) {
      var row = {};
      studentHeaders.forEach(function(header, index) { if (header) row[header] = values[index]; });
      return row;
    });
  }

  var sourceRows = students.map(function(row) {
    return {
      '학생ID':String(wmLmsCell_(row, ['학생ID','Student_ID','studentId']) || '').trim(),
      '학생이름':String(wmLmsCell_(row, ['학생이름','이름','Student_Name','studentName']) || '').trim()
    };
  }).filter(function(row) { return !!row['학생ID']; });

  var headers = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getDisplayValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  var missing = ['학생ID','학생이름'].filter(function(header) { return headers.indexOf(header) < 0; });
  if (missing.length) throw new Error(settingSheet.getName() + ' 필수 컬럼 누락: ' + missing.join(', '));

  var idColumn = headers.indexOf('학생ID') + 1;
  var existingById = {};
  if (settingSheet.getLastRow() >= 2) {
    settingSheet.getRange(2, idColumn, settingSheet.getLastRow() - 1, 1).getDisplayValues().forEach(function(row, index) {
      var key = String(row[0] || '').trim().toUpperCase();
      if (key && !existingById[key]) existingById[key] = index + 2;
    });
  }

  var updated = 0;
  var added = 0;
  sourceRows.forEach(function(source) {
    var key = source['학생ID'].toUpperCase();
    var rowNumber = existingById[key];
    if (!rowNumber) {
      rowNumber = Math.max(settingSheet.getLastRow() + 1, 2);
      existingById[key] = rowNumber;
      added++;
    } else {
      updated++;
    }
    settingSheet.getRange(rowNumber, headers.indexOf('학생ID') + 1).setValue(source['학생ID']);
    settingSheet.getRange(rowNumber, headers.indexOf('학생이름') + 1).setValue(source['학생이름']);
  });

  SpreadsheetApp.flush();
  return {
    success:true,
    log:{sheetName:logSheet.getName(), headers:logHeaders, preservedRows:Math.max(logSheet.getLastRow() - 1, 0)},
    setting:{sheetName:settingSheet.getName(), rows:sourceRows.length, updated:updated, added:added},
    message:'5-1 최종 14컬럼 적용 및 5-2 학생 기본정보 동기화 완료'
  };
}

function wmGetSendCenterDataForLms_(e) {
  try {
    /* WM_SEND_CENTER_TAB_LAZY_LOAD_V1
     * 5-1과 5-2를 동시에 읽지 않고 현재 선택 탭만 읽습니다.
     * 5-1은 발송내역만, 5-2는 자동발송설정과 학생 선택목록만 반환합니다. */
    var p = (e && e.parameter) ? e.parameter : {};
    var tab = String(p.tab || 'log').trim().toLowerCase();
    var actor = wmGetLmsRequestActor_(e);
    var restricted = actor && actor.found && /^(TEACHER|SUB_LEADER)$/.test(wmNormalizeLmsRole_(actor.role));

    if (tab === 'setting') {
      var settingResult = wmReadFirstExistingLmsSheetAsObjects_(['5-2.자동발송설정_DB']);
      var settingRows = settingResult.rows || [];
      var studentsResult = wmReadLmsSheetAsObjects_('1.학생관리_DB');
      var targetStudents = studentsResult.rows || [];
      if (restricted) {
        targetStudents = wmFilterStudentsForActor_(targetStudents, actor);
        var allowedSettingIds = wmStudentIdSet_(targetStudents);
        settingRows = settingRows.filter(function(row){
          var studentId = String(wmLmsCell_(row, ['학생ID','Student_ID','studentId']) || '').trim().toUpperCase();
          return !!allowedSettingIds[studentId];
        });
      }
      return {
        success:true,
        tab:'setting',
        settingSheetName:settingResult.sheetName || '5-2.자동발송설정_DB',
        settingHeaders:settingResult.headers || [],
        settingRows:settingRows,
        targetStudents:targetStudents.map(function(row){
          return {
            studentId:String(wmLmsCell_(row, ['학생ID','Student_ID','studentId']) || '').trim(),
            studentName:String(wmLmsCell_(row, ['학생이름','이름','Student_Name']) || '').trim()
          };
        }).filter(function(row){ return !!row.studentId; })
      };
    }

    var logResult = wmReadFirstExistingLmsSheetAsObjects_(['5-1.발송내역_DB']);
    var logRows = logResult.rows || [];
    if (restricted) {
      var allowedStudents = wmFilterStudentsForActor_(wmReadLmsSheetAsObjects_('1.학생관리_DB').rows || [], actor);
      var allowedTeacherNames = {};
      allowedStudents.forEach(function(row){
        var teacherName = String(wmLmsCell_(row, ['교사명','담당교사','선생님','teacherName']) || '').trim().toUpperCase();
        if (teacherName) allowedTeacherNames[teacherName] = true;
      });
      logRows = logRows.filter(function(row){
        var teacherName = String(wmLmsCell_(row, ['교사명','담당교사','선생님','teacherName']) || '').trim().toUpperCase();
        return !!allowedTeacherNames[teacherName];
      });
    }
    return {
      success:true,
      tab:'log',
      logSheetName:logResult.sheetName || '5-1.발송내역_DB',
      logHeaders:logResult.headers || [],
      logRows:logRows
    };
  } catch (err) {
    return {
      success:false,
      message:'발송센터 DB 조회 오류',
      error:String(err && err.message ? err.message : err),
      logRows:[],
      settingRows:[]
    };
  }
}

function wmSendSettingHeaders_() {
  return ['학생ID','학생이름','리포트유형','발송주기','발송요일','발송일','발송시간','자동발송상태','비고','등록일시','수정일시'];
}

function wmSendSettingTimestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/* 5-2.자동발송설정_DB 신규등록·수정 전용 API입니다. API 실제 발송은 수행하지 않습니다. */
function wmSaveSendSettingForLms_(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var actor = wmGetLmsRequestActor_(e);
    if (!actor || !actor.found) return {success:false, message:'로그인 사용자 정보를 확인할 수 없습니다.'};
    var ss = getLmsSpreadsheet_();
    var sheet = ss.getSheetByName('5-2.자동발송설정_DB');
    if (!sheet) return {success:false, message:'5-2.자동발송설정_DB 시트를 찾을 수 없습니다.'};
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
    var required = wmSendSettingHeaders_();
    var missing = required.filter(function(header){ return headers.indexOf(header) < 0; });
    if (missing.length) return {success:false, message:'5-2.자동발송설정_DB 필수 컬럼 누락: ' + missing.join(', ')};

    var studentId = String(p.studentId || p.targetId || p['학생ID'] || '').trim();
    var studentName = String(p.studentName || p.targetName || p['학생이름'] || '').trim();
    if (!studentId) return {success:false, message:'학생ID를 확인하세요.'};

    var studentRows = wmFilterStudentsForActor_(wmReadLmsSheetAsObjects_('1.학생관리_DB').rows || [], actor);
    var matchedStudent = null;
    studentRows.some(function(row){
      var rowId = String(wmLmsCell_(row, ['학생ID','Student_ID','studentId']) || '').trim();
      if (rowId.toUpperCase() !== studentId.toUpperCase()) return false;
      matchedStudent = row;
      return true;
    });
    if (!matchedStudent) return {success:false, message:'학생관리_DB에서 학생ID 또는 권한을 확인할 수 없습니다.'};
    studentName = String(wmLmsCell_(matchedStudent, ['학생이름','이름','Student_Name']) || studentName).trim();

    var rowNumber = 0;
    if (sheet.getLastRow() >= 2) {
      var idIndex = headers.indexOf('학생ID');
      var ids = sheet.getRange(2, idIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0] || '').trim().toUpperCase() === studentId.toUpperCase()) { rowNumber = i + 2; break; }
      }
    }
    var now = wmSendSettingTimestamp_();
    var existing = rowNumber ? sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0] : [];
    var values = {
      '학생ID':studentId, '학생이름':studentName,
      '리포트유형':String(p.reportType || '').trim(), '발송주기':String(p.sendCycle || '').trim(),
      '발송요일':String(p.sendWeekday || '').trim(), '발송일':String(p.sendDay || '').trim(), '발송시간':String(p.sendTime || '').trim(),
      '자동발송상태':String(p.autoStatus || '').trim(),
      '비고':String(p.note || '').trim(), '등록일시':rowNumber ? String(existing[headers.indexOf('등록일시')] || '').trim() : now, '수정일시':now
    };
    if (!rowNumber) return {success:false, message:'학생관리_DB 상시 참조 행을 찾을 수 없습니다: ' + studentId};
    var writableHeaders = headers.slice(2);
    var output = writableHeaders.map(function(header){
      return values.hasOwnProperty(header) ? values[header] : existing[headers.indexOf(header)];
    });
    sheet.getRange(rowNumber, 3, 1, writableHeaders.length).setValues([output]);
    SpreadsheetApp.flush();
    var savedValues = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
    var savedRow = {};
    headers.forEach(function(header, index){ savedRow[header] = savedValues[index]; });
    if (String(savedRow['학생ID'] || '').trim().toUpperCase() !== studentId.toUpperCase()) return {success:false, message:'구글시트 저장 후 학생ID 재확인에 실패했습니다.'};
    return {success:true, sheetName:'5-2.자동발송설정_DB', savedRow:savedRow, studentId:studentId, rowNumber:rowNumber, mode:existing.length ? 'update' : 'insert', message:existing.length ? '구글시트 자동발송설정을 수정했습니다.' : '구글시트 자동발송설정을 등록했습니다.'};
  } catch (err) {
    return {success:false, message:'자동발송설정 저장 오류', error:String(err && err.message ? err.message : err)};
  }
}

/* 기존 6.성적등급_DB 전체 행을 확정된 성적표보기·최근성적표보기 월 누적 공식으로 즉시 갱신합니다. */
function wmRunGradeMonthlyCumulativeRefresh() {
  wmSyncLearningDatesToSheets_();
  return wmSyncGradePrimaryDataAfterLearningSave_(getLmsSpreadsheet_());
}

/* =========================================================
 * WM_REPORT_SCHEDULE_TRIGGER_V2
 * [성적표 생성·발행 확정 기준]
 *
 * 1. Weekly Report
 *    - 매주 월요일 00:01(Asia/Seoul)에 생성합니다.
 *    - 해당 월 1일부터 직전 종료 주차 말일까지 2.학습기록_DB 누적기록을 한 번 집계합니다.
 *    - 진행 중인 세트는 제외하며, 완료된 다음 기간에 반영합니다.
 *
 * 2. 월의 마지막 주
 *    - 마지막 주가 4주차든 5주차든 성적표 내부 주차 기록은 생성합니다.
 *    - 마지막 주 집계 종료일은 일요일이 아니라 해당 월의 말일입니다.
 *    - 마지막 주 Weekly Report는 별도로 생성·발행하지 않습니다.
 *    - 월말까지 완료된 세트는 다음 달 1일 Monthly Report에 합산합니다.
 *
 * 3. Monthly Report
 *    - 매월 1일 00:01(Asia/Seoul)에 생성합니다.
 *    - 직전 월 1일~말일까지 2.학습기록_DB 누적기록을 한 번 집계합니다.
 *    - 해당 월의 모든 주차 기록을 포함합니다.
 *
 * 4. 6주차 처리
 *    - 성적표에는 6주차를 생성하지 않습니다.
 *    - 기존 5주차 기간을 해당 월 말일까지 확장합니다.
 *    - 달력상 5주차와 6주차 데이터는 모두 5주차 기록으로 통합합니다.
 *    - 5주차 Weekly Report는 발행하지 않고 다음 달 Monthly Report로 발행합니다.
 *
 * 5. 예약 실행
 *    - 월요일 00:01에는 Weekly 생성 대상을 처리합니다.
 *    - 매월 1일 00:01에는 Monthly 생성 대상을 처리합니다.
 *    - 매월 1일이 월요일인 경우 한 번의 동기화에서 Weekly와 Monthly를 함께 처리합니다.
 *    - 그 외 날짜와 시각에는 성적표 생성 동기화를 실행하지 않습니다.
 *    - 설치는 wmInstallReportScheduleTrigger_()를 테스트부스에서 1회 실행합니다.
 * ========================================================= */
function wmScheduledReportGenerationTick_() {
  var now = new Date();
  var nowTime = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  var dayOfWeek = Number(Utilities.formatDate(now, 'Asia/Seoul', 'u')); // 월=1, 일=7
  var dayOfMonth = Number(Utilities.formatDate(now, 'Asia/Seoul', 'd'));
  var weeklyDue = dayOfWeek === 1;
  var monthlyDue = dayOfMonth === 1;

  if (nowTime < '00:01' || (!weeklyDue && !monthlyDue)) {
    return {success:true, skipped:true, message:'매주 월요일 또는 매월 1일 00:01 이후 발행 대기'};
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return {success:true, skipped:true, message:'중복 실행 차단'};
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var dateKey = wmReportSeoulDateKey_(now);
    var runKey = dateKey + '|REPORT_0001';

    if (props.getProperty('WM_REPORT_LAST_SCHEDULE_RUN') === runKey) {
      return {success:true, skipped:true, message:'오늘 예약 성적표 생성 완료'};
    }

    var reportResult = wmInitializeReportsFromLearningRecordsForLms_();
    if (!reportResult || reportResult.success !== true) {
      return {success:false, message:'4-1.성적표생성_DB 생성 실패', report:reportResult || null};
    }

    var gradeResult = wmSyncGradePrimaryDataAfterLearningSave_(getLmsSpreadsheet_());
    if (!gradeResult || gradeResult.success !== true) {
      return {success:false, message:'6.성적등급_DB 생성 실패', report:reportResult, grade:gradeResult || null};
    }

    props.setProperty('WM_REPORT_LAST_SCHEDULE_RUN', runKey);
    return {
      success:true,
      scheduleDate:dateKey,
      report:reportResult,
      grade:gradeResult,
      message:'4-1.성적표생성_DB → 6.성적등급_DB 월요일/매월 1일 00:01 생성 완료'
    };
  } finally {
    lock.releaseLock();
  }
}
function wmInstallReportScheduleTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'wmScheduledReportGenerationTick_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('wmScheduledReportGenerationTick_')
    .timeBased()
    .everyMinutes(1)
    .create();
  return {
    success:true,
    message:'매주 월요일/매월 1일 00:01 성적표생성_DB → 성적등급_DB 자동생성 트리거 설치 완료'
  };
}

function wmInstallReportScheduleTrigger() {
  return wmInstallReportScheduleTrigger_();
}

function wmNormalizeReportGenerationSetIds() {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('4-1.성적표생성_DB');
  if (!sheet) return {success:false, message:'4-1.성적표생성_DB 시트를 찾을 수 없습니다.'};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {success:true, changedCount:0, message:'변경할 Set_ID가 없습니다.'};

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  var setCol = headers.indexOf('Set_ID') + 1;
  if (setCol < 1) return {success:false, message:'4-1.성적표생성_DB Set_ID 컬럼을 찾을 수 없습니다.'};

  var range = sheet.getRange(2, setCol, lastRow - 1, 1);
  var values = range.getDisplayValues();
  var changedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var beforeValue = String(values[i][0] || '').trim();
    var afterValue = wmReportStoredSetId_(beforeValue);
    if (afterValue && afterValue !== beforeValue) {
      values[i][0] = afterValue;
      changedCount += 1;
    }
  }

  range.setNumberFormat('@');
  if (changedCount > 0) range.setValues(values);

  return {
    success:true,
    changedCount:changedCount,
    message:'4-1.성적표생성_DB Set_ID를 WM3-1-5 형식으로 정규화 완료'
  };
}

function wmRecoverReportDatabases() {
  var reportResult = wmInitializeReportsFromLearningRecordsForLms_();
  Logger.log('WM_REPORT_RECOVERY_RESULT=' + JSON.stringify(reportResult || null));
  if (!reportResult || reportResult.success !== true) return reportResult;

  var gradeResult = wmSyncGradePrimaryDataAfterLearningSave_(getLmsSpreadsheet_());
  if (!gradeResult || gradeResult.success !== true) return gradeResult;

  return {
    success:true,
    report:reportResult,
    grade:gradeResult,
    message:'4-1.성적표생성_DB → 6.성적등급_DB 기존 누락 복구 완료'
  };
}

/* =========================================================
 * WM_GRADE_DB_SCHEDULE_ONLY_V1
 * 성적표 생성은 wmScheduledReportGenerationTick_() 예약 함수에서만 실행합니다.
 * 시트 열기(onOpen) 또는 직접 수정(onEdit)으로 성적표가 임의 생성되지 않도록 차단합니다.
 * ========================================================= */
function onOpen(e) {
  try {
    wmSyncSendCenterPrimaryDataFromStudents_(null, e && e.source ? e.source : SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    console.error('발송센터 구글시트 직접 저장 오류: ' + String(err && err.message ? err.message : err));
  }
}

function wmStudentMasterDirectEditSyncTrigger_(e) {
  /* WM_STUDENT_MASTER_DIRECT_EDIT_SYNC_V2_20260821
   * Google Sheet에서 직접 수정하는 경우를 위한 실제 편집 트리거 본체입니다.
   * 독립형 Apps Script에서는 단순 onEdit가 대상 스프레드시트에 자동 연결되지 않으므로
   * 설치형 onEdit 트리거도 같은 본체를 사용합니다. */
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== '1.학생관리_DB') return;
    if (e.range.getRow() < 2) return;

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h){
      return String(h || '').trim();
    });
    var watched = ['학생이름','학교','학년','반명','Class','반','교사명','담당교사'];
    var startCol = e.range.getColumn();
    var endCol = startCol + e.range.getNumColumns() - 1;
        var idxStudentSerial = headers.indexOf('학생일련번호');
    var singleCell =
      e.range.getNumRows() === 1 &&
      e.range.getNumColumns() === 1;

    if (
      singleCell &&
      idxStudentSerial >= 0 &&
      startCol === idxStudentSerial + 1
    ) {
      wmRepairEditedStudentSerialFast_(
        sheet,
        headers,
        e.range.getRow()
      );
      return;
    }

    var idxParentPhone = headers.indexOf('학부모연락처'); if (idxParentPhone >= 0 && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && startCol === idxParentPhone + 1) { var phone = String(e.value || '').replace(/\D/g,''); if (phone.length === 10 && phone.charAt(0) === '1') phone = '0' + phone; if (/^010\d{8}$/.test(phone)) e.range.setNumberFormat('@').setValue(phone.replace(/^(\d{3})(\d{4})(\d{4})$/,'$1-$2-$3')); return; }
    var idxParentPhone = headers.indexOf('학부모연락처');
    var touchesStudentInfo = watched.some(function(header){
      var idx = headers.indexOf(header);
      return idx >= 0 && (idx + 1) >= startCol && (idx + 1) <= endCol;
    });
    if (!touchesStudentInfo) return;

    var idxStudentId = headers.indexOf('학생ID');
    if (idxStudentId < 0) return;

    var studentIds = [];
    var startRow = Math.max(2, e.range.getRow());
    var endRow = e.range.getRow() + e.range.getNumRows() - 1;
    for (var rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
      var studentId = String(sheet.getRange(rowNumber, idxStudentId + 1).getDisplayValue() || '').trim();
      if (studentId) studentIds.push(studentId);
    }
    if (!studentIds.length) return;
    return wmSyncStudentMasterInfoToCurrentProgressBatch_(e.source || sheet.getParent(), studentIds);
  } catch (err) {
    console.error('학생관리_DB → 현재진행_DB 기본정보 동기화 오류: ' + String(err && err.message ? err.message : err));
  }
}

function onEdit(e) {
  return wmStudentMasterDirectEditSyncTrigger_(e);
}

/* 독립형 REAL Code.gs가 1.학생관리_DB 직접 편집도 즉시 잡도록 설치형 onEdit를 1회 등록합니다.
 * 이미 같은 핸들러가 있으면 중복 생성하지 않습니다. */
function wmInstallStudentMasterDirectEditSyncTrigger() {
  return wmInstallStudentMasterDirectEditSyncTrigger_();
}
function wmInstallStudentMasterDirectEditSyncTrigger_() {
  var ss = getLmsSpreadsheet_();
  var phoneSheet = ss.getSheetByName('1.학생관리_DB'), phoneHeaders = phoneSheet ? phoneSheet.getRange(1,1,1,phoneSheet.getLastColumn()).getDisplayValues()[0] : [], phoneCol = phoneHeaders.indexOf('학부모연락처'); if (phoneSheet && phoneCol >= 0) phoneSheet.getRange(2, phoneCol + 1, Math.max(phoneSheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  var handler = 'wmStudentMasterDirectEditSyncTrigger_';
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) found = true;
  });

  if (!found) {
    ScriptApp.newTrigger(handler)
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }

  return {
    success:true,
    spreadsheetId:ss.getId(),
    handler:handler,
    installed:!found,
    alreadyInstalled:found
  };
}


/* WM_STUDENT_INFO_CURRENT_PROGRESS_INTEGRATED_SETUP_V1_20260821
 * 한 번 실행으로 다음 3가지를 함께 완료합니다.
 * 1) 기존 8.현재진행_DB 전체 학생정보를 1.학생관리_DB 기준으로 일괄 정합화
 * 2) 이후 LMS 저장경로의 즉시동기화는 기존 연결을 그대로 사용
 * 3) 구글시트 직접수정용 설치형 onEdit 트리거를 중복 없이 1회 설치
 * 과거 2.학습기록_DB 및 현재진행의 진행/점수 컬럼은 수정하지 않습니다. */
function wmSyncAllExistingStudentInfoToCurrentProgress_() {
  var ss = getLmsSpreadsheet_();
  var studentSheet = ss.getSheetByName('1.학생관리_DB');
  var progressSheet = ss.getSheetByName('8.현재진행_DB');

  if (!studentSheet || !progressSheet) {
    return {success:false, message:'1.학생관리_DB 또는 8.현재진행_DB 시트를 찾을 수 없습니다.'};
  }
  if (studentSheet.getLastRow() < 2 || progressSheet.getLastRow() < 2) {
    return {success:true, updatedStudents:0, updatedCells:0, message:'동기화 대상 행 없음'};
  }

  var studentLastCol = studentSheet.getLastColumn();
  var progressLastCol = progressSheet.getLastColumn();
  var studentHeaders = studentSheet.getRange(1, 1, 1, studentLastCol).getDisplayValues()[0].map(function(v){
    return String(v || '').trim();
  });
  var progressHeaders = progressSheet.getRange(1, 1, 1, progressLastCol).getDisplayValues()[0].map(function(v){
    return String(v || '').trim();
  });

  var studentIdCol = studentHeaders.indexOf('학생ID');
  var progressIdCol = progressHeaders.indexOf('학생ID');
  if (studentIdCol < 0 || progressIdCol < 0) {
    return {success:false, message:'학생ID 컬럼을 확인하세요.'};
  }

  function findHeaderIndex_(headers, names) {
    for (var i = 0; i < names.length; i++) {
      var idx = headers.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  var studentCols = {
    학생이름: findHeaderIndex_(studentHeaders, ['학생이름','학생명','이름']),
    학교: findHeaderIndex_(studentHeaders, ['학교','학교명']),
    학년: findHeaderIndex_(studentHeaders, ['학년']),
    Class: findHeaderIndex_(studentHeaders, ['반명','Class','반']),
    교사명: findHeaderIndex_(studentHeaders, ['교사명','담당교사'])
  };
  var progressCols = {
    학생이름: progressHeaders.indexOf('학생이름'),
    학교: progressHeaders.indexOf('학교'),
    학년: progressHeaders.indexOf('학년'),
    Class: progressHeaders.indexOf('Class'),
    교사명: progressHeaders.indexOf('교사명')
  };

  var studentRows = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, studentLastCol).getDisplayValues();
  var progressRowCount = progressSheet.getLastRow() - 1;
  var progressRows = progressSheet.getRange(2, 1, progressRowCount, progressLastCol).getValues();

  var studentMap = {};
  studentRows.forEach(function(row) {
    var id = String(row[studentIdCol] || '').trim();
    var key = id.toUpperCase();
    if (!key) return;
    studentMap[key] = {
      학생이름: studentCols.학생이름 >= 0 ? String(row[studentCols.학생이름] || '').trim() : '',
      학교: studentCols.학교 >= 0 ? String(row[studentCols.학교] || '').trim() : '',
      학년: studentCols.학년 >= 0 ? String(row[studentCols.학년] || '').trim() : '',
      Class: studentCols.Class >= 0 ? String(row[studentCols.Class] || '').trim() : '',
      교사명: studentCols.교사명 >= 0 ? String(row[studentCols.교사명] || '').trim() : ''
    };
  });

  var updatedStudents = {};
  var updatedCells = 0;
  var changedColumns = {};

  progressRows.forEach(function(row) {
    var id = String(row[progressIdCol] || '').trim();
    var key = id.toUpperCase();
    var source = studentMap[key];
    if (!key || !source) return;

    Object.keys(progressCols).forEach(function(name) {
      var col = progressCols[name];
      if (col < 0) return;
      var nextValue = String(source[name] || '').trim();
      var currentValue = String(row[col] || '').trim();
      if (currentValue === nextValue) return;
      row[col] = nextValue;
      changedColumns[col] = true;
      updatedCells++;
      updatedStudents[key] = id;
    });
  });

  Object.keys(changedColumns).forEach(function(colText) {
    var col = Number(colText);
    var columnValues = progressRows.map(function(row){ return [row[col]]; });
    progressSheet.getRange(2, col + 1, progressRowCount, 1).setValues(columnValues);
  });

  var changedIds = Object.keys(updatedStudents).map(function(key){ return updatedStudents[key]; });
  if (changedIds.length) wmClearRuntimeCachesForStudents_(changedIds);

  return {
    success:true,
    updatedStudents:changedIds.length,
    updatedCells:updatedCells,
    message:'기존 현재진행_DB 학생정보 전체 동기화 완료'
  };
}

function wmApplyStudentInfoCurrentProgressSyncAll() {
  var fullSync = wmSyncAllExistingStudentInfoToCurrentProgress_();
  if (!fullSync || fullSync.success !== true) {
    return {success:false, stage:'FULL_SYNC', fullSync:fullSync};
  }

  var trigger = wmInstallStudentMasterDirectEditSyncTrigger_();
  return {
    success:true,
    fullSync:fullSync,
    lmsImmediateSync:true,
    directSheetEditSync:true,
    trigger:trigger,
    message:'기존 전체동기화 + 향후 LMS 즉시동기화 + 구글시트 직접수정 자동동기화 설정 완료'
  };
}

/* =========================================================
   🔒 WM CODE.GS LMS PHASE 1 CORE OPERATIONS LOCK V1
   대상: 대시보드 + 학생관리 + 교사관리 + 반관리
   실제 서버 함수 참조·권한 필터·최소조회·저장 경로를 동결합니다.
   ========================================================= */
const WM_CODEGS_LMS_PHASE1_CORE_LOCK_V1 = Object.freeze({
  version: 'WM_CODEGS_LMS_PHASE1_CORE_LOCK_V1.0',
  date: '2026-08-06',
  DASHBOARD: Object.freeze({
    load: getDashboardForLmsApi_,
    summary: wmBuildDashboardSummary_,
    actorFilter: wmFilterStudentsForActor_,
    cacheGet: wmCacheGetJson_,
    cachePut: wmCachePutJson_
  }),
  STUDENT_MANAGEMENT: Object.freeze({
    load: getStudentsForLmsApi_,
    save: saveStudentLearningModeApi_,
    actorFilter: wmFilterStudentsForActor_,
    currentSet: buildCurrentProgressSetMapForLms_,
    clearStudentCache: wmClearRuntimeCachesForStudent_
  }),
  TEACHER_MANAGEMENT: Object.freeze({
    load: wmGetSettingCenterForLms_,
    checkTeacherId: wmCheckSettingTeacherIdForLms_,
    saveTeacher: wmSaveSettingTeacherForLms_,
    saveTeacherStatus: wmSaveSettingTeacherStatusForLms_
  }),
  CLASS_MANAGEMENT: Object.freeze({
    load: wmGetSettingCenterForLms_,
    saveClass: wmSaveSettingClassForLms_,
    actorFilter: wmFilterStudentsForActor_,
    clearStudentsCache: wmClearRuntimeCachesForStudents_
  })
});

function wmGetCodeGsLmsPhase1CoreLockStatus_() {
  var registry = WM_CODEGS_LMS_PHASE1_CORE_LOCK_V1;
  var expected = {
    DASHBOARD: {
      load: getDashboardForLmsApi_, summary: wmBuildDashboardSummary_,
      actorFilter: wmFilterStudentsForActor_, cacheGet: wmCacheGetJson_, cachePut: wmCachePutJson_
    },
    STUDENT_MANAGEMENT: {
      load: getStudentsForLmsApi_, save: saveStudentLearningModeApi_,
      actorFilter: wmFilterStudentsForActor_, currentSet: buildCurrentProgressSetMapForLms_,
      clearStudentCache: wmClearRuntimeCachesForStudent_
    },
    TEACHER_MANAGEMENT: {
      load: wmGetSettingCenterForLms_, checkTeacherId: wmCheckSettingTeacherIdForLms_,
      saveTeacher: wmSaveSettingTeacherForLms_, saveTeacherStatus: wmSaveSettingTeacherStatusForLms_
    },
    CLASS_MANAGEMENT: {
      load: wmGetSettingCenterForLms_, saveClass: wmSaveSettingClassForLms_,
      actorFilter: wmFilterStudentsForActor_, clearStudentsCache: wmClearRuntimeCachesForStudents_
    }
  };
  var modules = {};
  var allVerified = Object.isFrozen(registry);
  Object.keys(expected).forEach(function(moduleName) {
    var group = registry[moduleName] || {};
    var checks = {};
    var verified = Object.isFrozen(group);
    Object.keys(expected[moduleName]).forEach(function(functionName) {
      checks[functionName] = typeof group[functionName] === 'function' &&
        group[functionName] === expected[moduleName][functionName];
      verified = verified && checks[functionName];
    });
    modules[moduleName] = {verified: verified, functions: checks};
    allVerified = allVerified && verified;
  });
  return {
    version: registry.version,
    date: registry.date,
    registryFrozen: Object.isFrozen(registry),
    nestedRegistryFrozen: Object.keys(expected).every(function(name){ return Object.isFrozen(registry[name]); }),
    modules: modules,
    allVerified: allVerified
  };
}

function wmIsCodeGsLmsPhase1ProtectedRequest_(action, mode) {
  var key = String(action || mode || '').trim();
  return [
    'dashboard','students','saveStudentLearningMode',
    'getSettingCenter','checkTeacherId','saveSettingTeacher','saveSettingTeacherStatus','saveSettingClass'
  ].indexOf(key) !== -1;
}
/* 🔒 WM CODE.GS LMS PHASE 1 CORE OPERATIONS LOCK V1 END */

/* =========================================================
   🔒 WM CODE.GS LMS PHASE 2 REPORT / GRADE LOCK V1
   대상: 성적표생성 + 성적등급 + 최근성적표
   실제 서버 함수·조회·산정·자동생성 경로를 동결합니다.
   ========================================================= */
const WM_CODEGS_LMS_PHASE2_REPORT_GRADE_LOCK_V1 = Object.freeze({
  version: 'WM_CODEGS_LMS_PHASE2_REPORT_GRADE_LOCK_V1.0',
  date: '2026-08-06',
  REPORT_GENERATION: Object.freeze({
    load: wmGetReportsForLmsApi_,
    initialize: wmInitializeReportsFromLearningRecordsForLms_,
    monthlyData: wmGetMonthlyReportDataForLms_,
    scheduleRun: wmScheduledReportGenerationTick_,
    scheduleInstall: wmInstallReportScheduleTrigger_,
    primarySync: wmSyncGradePrimaryDataAfterLearningSave_
  }),
  GRADE_LEVEL: Object.freeze({
    load: wmGetGradeLevelsForLmsApi_,
    monthlyCalculate: wmCalculateMonthlyCumulativeGrade_,
    autoComment: wmBuildGradeAutoComment_,
    fullRefresh: wmRunGradeMonthlyCumulativeRefresh
  }),
  RECENT_REPORT: Object.freeze({
    months: wmGetRecentReportMonthsForLms_,
    monthlyData: wmGetMonthlyReportDataForLms_
  }),
  RULES: Object.freeze({
    weeklyMonthly: true,
    learningStartEnd: true,
    monthlyCumulative: true,
    lastWeekMonthly: true,
    partialWeekProration: true,
    fullAbsenceZero: true,
    completedSetOnly: true,
    duplicatePrevention: true,
    fourMonthFixed: true,
    currentMonthProgressive: true,
    previousThreeMonths: true,
    sameMonthlyFormula: true,
    studentIndependentQuery: true
  })
});

function wmGetCodeGsLmsPhase2ReportGradeLockStatus_() {
  var registry=WM_CODEGS_LMS_PHASE2_REPORT_GRADE_LOCK_V1;
  var expected={
    REPORT_GENERATION:{
      load:wmGetReportsForLmsApi_,
      initialize:wmInitializeReportsFromLearningRecordsForLms_,
      monthlyData:wmGetMonthlyReportDataForLms_,
      scheduleRun:wmScheduledReportGenerationTick_,
      scheduleInstall:wmInstallReportScheduleTrigger_,
      primarySync:wmSyncGradePrimaryDataAfterLearningSave_
    },
    GRADE_LEVEL:{
      load:wmGetGradeLevelsForLmsApi_,
      monthlyCalculate:wmCalculateMonthlyCumulativeGrade_,
      autoComment:wmBuildGradeAutoComment_,
      fullRefresh:wmRunGradeMonthlyCumulativeRefresh
    },
    RECENT_REPORT:{
      months:wmGetRecentReportMonthsForLms_,
      monthlyData:wmGetMonthlyReportDataForLms_
    }
  };
  var modules={};
  var allVerified=Object.isFrozen(registry)&&Object.isFrozen(registry.RULES);

  Object.keys(expected).forEach(function(moduleName){
    var group=registry[moduleName]||{};
    var checks={};
    var verified=Object.isFrozen(group);
    Object.keys(expected[moduleName]).forEach(function(functionName){
      checks[functionName]=typeof group[functionName]==='function'&&
        group[functionName]===expected[moduleName][functionName];
      verified=verified&&checks[functionName];
    });
    modules[moduleName]={verified:verified,functions:checks};
    allVerified=allVerified&&verified;
  });

  var rulesVerified=Object.keys(registry.RULES).every(function(ruleName){
    return registry.RULES[ruleName]===true;
  });
  allVerified=allVerified&&rulesVerified;

  return {
    version:registry.version,
    date:registry.date,
    registryFrozen:Object.isFrozen(registry),
    nestedRegistryFrozen:Object.keys(expected).every(function(name){
      return Object.isFrozen(registry[name]);
    }),
    rulesFrozen:Object.isFrozen(registry.RULES),
    rulesVerified:rulesVerified,
    modules:modules,
    allVerified:allVerified
  };
}

function wmIsCodeGsLmsPhase2ProtectedRequest_(action,mode) {
  var key=String(action||mode||'').trim();
  return [
    'reports','monthlyReportData','recentReportMonths',
    'initializeReportsFromLearningRecords','gradeLevels','gradeLevel'
  ].indexOf(key)!==-1;
}
/* 🔒 WM CODE.GS LMS PHASE 2 REPORT / GRADE LOCK V1 END */

/* =========================================================
   🔒 WM CODE.GS LMS PHASE 3 PROFILE / SEND CENTER LOCK V1
   대상: 프로필센터 + 발송센터
   학생별 독립조회·권한필터·발송내역/자동발송설정 서버 경로 동결
   ========================================================= */
const WM_CODEGS_LMS_PHASE3_PROFILE_SEND_LOCK_V1 = Object.freeze({
  version:'WM_CODEGS_LMS_PHASE3_PROFILE_SEND_LOCK_V1.0',
  date:'2026-08-06',

  PROFILE_CENTER:Object.freeze({
    students:getStudentsForLmsApi_,
    records:getLearningRecordsForLmsApi_,
    reports:wmGetReportsForLmsApi_,
    grades:wmGetGradeLevelsForLmsApi_,
    studentScope:wmFilterStudentsForActor_,
    rowScope:wmFilterRowsByStudentSet_
  }),

  SEND_CENTER:Object.freeze({
    load:wmGetSendCenterDataForLms_,
    saveSetting:wmSaveSendSettingForLms_,
    studentScope:wmFilterStudentsForActor_
  }),

  RULES:Object.freeze({
    studentIndependentQuery:true,
    selectedStudentIsolation:true,
    roleScope:true,
    sectionIndependentLoad:true,
    profilePagination:true,
    sendTabIndependentLoad:true,
    sendLogAndSettingSeparated:true,
    studentNameIdLink:true,
    parentRecipientData:true,
    saveAndEdit:true,
    profileSendHistory:true
  })
});

function wmGetCodeGsLmsPhase3ProfileSendLockStatus_(){
  var registry=WM_CODEGS_LMS_PHASE3_PROFILE_SEND_LOCK_V1;

  var expected={
    PROFILE_CENTER:{
      students:getStudentsForLmsApi_,
      records:getLearningRecordsForLmsApi_,
      reports:wmGetReportsForLmsApi_,
      grades:wmGetGradeLevelsForLmsApi_,
      studentScope:wmFilterStudentsForActor_,
      rowScope:wmFilterRowsByStudentSet_
    },
    SEND_CENTER:{
      load:wmGetSendCenterDataForLms_,
      saveSetting:wmSaveSendSettingForLms_,
      studentScope:wmFilterStudentsForActor_
    }
  };

  var modules={};
  var allVerified=
    Object.isFrozen(registry) &&
    Object.isFrozen(registry.RULES);

  Object.keys(expected).forEach(function(moduleName){
    var group=registry[moduleName]||{};
    var checks={};
    var verified=Object.isFrozen(group);

    Object.keys(expected[moduleName]).forEach(function(functionName){
      checks[functionName]=
        typeof group[functionName]==='function' &&
        group[functionName]===expected[moduleName][functionName];
      verified=verified&&checks[functionName];
    });

    modules[moduleName]={
      verified:verified,
      functions:checks
    };
    allVerified=allVerified&&verified;
  });

  var rulesVerified=Object.keys(registry.RULES).every(function(ruleName){
    return registry.RULES[ruleName]===true;
  });

  allVerified=allVerified&&rulesVerified;

  return {
    version:registry.version,
    date:registry.date,
    registryFrozen:Object.isFrozen(registry),
    nestedRegistryFrozen:
      Object.isFrozen(registry.PROFILE_CENTER) &&
      Object.isFrozen(registry.SEND_CENTER),
    rulesFrozen:Object.isFrozen(registry.RULES),
    rulesVerified:rulesVerified,
    modules:modules,
    allVerified:allVerified
  };
}

function wmIsCodeGsLmsPhase3ProtectedRequest_(action,mode){
  var key=String(action||mode||'').trim();
  return [
    'students',
    'records',
    'learningRecords',
    'reports',
    'gradeLevels',
    'gradeLevel',
    'sendCenterData',
    'saveSendSetting'
  ].indexOf(key)!==-1;
}
/* 🔒 WM CODE.GS LMS PHASE 3 PROFILE / SEND CENTER LOCK V1 END */

/* =========================================================
   🔒 WM CODE.GS LMS PHASE 4 COMMON ROLE / SPEED LOCK V1
   대상: 공통 권한 + 공통 속도
   서버 권한범위·최소조회·캐시·부분무효화·반복 전체조회 방지 경로 동결
   ========================================================= */
const WM_CODEGS_LMS_PHASE4_COMMON_ROLE_SPEED_LOCK_V1 = Object.freeze({
  version:'WM_CODEGS_LMS_PHASE4_COMMON_ROLE_SPEED_LOCK_V1.0',
  date:'2026-08-06',

  COMMON_ROLE:Object.freeze({
    actor:wmGetLmsRequestActor_,
    normalizeRole:wmNormalizeLmsRole_,
    subLeaderScope:wmSubLeaderScope_,
    filterStudents:wmFilterStudentsForActor_,
    filterRows:wmFilterRowsByStudentSet_
  }),

  COMMON_SPEED:Object.freeze({
    spreadsheet:getLmsSpreadsheet_,
    cacheGet:wmCacheGetJson_,
    cachePut:wmCachePutJson_,
    cacheRemove:wmCacheRemove_,
    cacheKey:wmCacheKey_,
    clearStudentCache:wmClearRuntimeCachesForStudent_,
    clearStudentsCache:wmClearRuntimeCachesForStudents_,
    initialPageCheck:wmIsLmsInitialPageRequest_,
    dashboard:getDashboardForLmsApi_,
    studentRows:wmGetStudentRowNumbersFast_,
    groupedRows:wmReadGroupedRowsFast_
  }),

  RULES:Object.freeze({
    fourRoles:true,
    menuAndDataScope:true,
    clientServerRoleMatch:true,
    firstScreenMinimum:true,
    loadAfterTabEntry:true,
    first15Or25Rows:true,
    serverSearchAndPaging:true,
    studentIndependentQuery:true,
    duplicateRequestPrevention:true,
    cacheReuse:true,
    partialCacheInvalidation:true,
    preventRepeatedFullDbRead:true
  })
});

function wmGetCodeGsLmsPhase4CommonRoleSpeedLockStatus_(){
  var registry=WM_CODEGS_LMS_PHASE4_COMMON_ROLE_SPEED_LOCK_V1;
  var expected={
    COMMON_ROLE:{
      actor:wmGetLmsRequestActor_,
      normalizeRole:wmNormalizeLmsRole_,
      subLeaderScope:wmSubLeaderScope_,
      filterStudents:wmFilterStudentsForActor_,
      filterRows:wmFilterRowsByStudentSet_
    },
    COMMON_SPEED:{
      spreadsheet:getLmsSpreadsheet_,
      cacheGet:wmCacheGetJson_,
      cachePut:wmCachePutJson_,
      cacheRemove:wmCacheRemove_,
      cacheKey:wmCacheKey_,
      clearStudentCache:wmClearRuntimeCachesForStudent_,
      clearStudentsCache:wmClearRuntimeCachesForStudents_,
      initialPageCheck:wmIsLmsInitialPageRequest_,
      dashboard:getDashboardForLmsApi_,
      studentRows:wmGetStudentRowNumbersFast_,
      groupedRows:wmReadGroupedRowsFast_
    }
  };

  var modules={};
  var allVerified=Object.isFrozen(registry)&&Object.isFrozen(registry.RULES);

  Object.keys(expected).forEach(function(moduleName){
    var group=registry[moduleName]||{};
    var checks={};
    var verified=Object.isFrozen(group);
    Object.keys(expected[moduleName]).forEach(function(key){
      checks[key]=typeof group[key]==='function'&&
        group[key]===expected[moduleName][key];
      verified=verified&&checks[key];
    });
    modules[moduleName]={verified:verified,functions:checks};
    allVerified=allVerified&&verified;
  });

  var rulesVerified=Object.keys(registry.RULES).every(function(key){
    return registry.RULES[key]===true;
  });
  allVerified=allVerified&&rulesVerified;

  return {
    version:registry.version,
    date:registry.date,
    registryFrozen:Object.isFrozen(registry),
    nestedRegistryFrozen:
      Object.isFrozen(registry.COMMON_ROLE)&&
      Object.isFrozen(registry.COMMON_SPEED),
    rulesFrozen:Object.isFrozen(registry.RULES),
    rulesVerified:rulesVerified,
    modules:modules,
    allVerified:allVerified
  };
}

function wmIsCodeGsLmsPhase4ProtectedRequest_(action,mode){
  var key=String(action||mode||'').trim();
  return [
    'dashboard','students','records','learningRecords',
    'reports','gradeLevels','gradeLevel',
    'getSettingCenter','sendCenterData'
  ].indexOf(key)!==-1;
}
/* 🔒 WM CODE.GS LMS PHASE 4 COMMON ROLE / SPEED LOCK V1 END */

/* =========================================================
   🔒 WM CODE.GS LMS PHASE 5 INTEGRATED ALL LOCK VERIFICATION V1
   1~4차 서버 Registry 동결·실제 함수 참조·모듈 결과 최종 통합 검증
   ========================================================= */
const WM_CODEGS_LMS_PHASE5_INTEGRATED_ALL_LOCK_V1 = Object.freeze({
  version:'WM_CODEGS_LMS_PHASE5_INTEGRATED_ALL_LOCK_V1.0',
  date:'2026-08-06',
  PHASE1:Object.freeze({
    registry:WM_CODEGS_LMS_PHASE1_CORE_LOCK_V1,
    status:wmGetCodeGsLmsPhase1CoreLockStatus_
  }),
  PHASE2:Object.freeze({
    registry:WM_CODEGS_LMS_PHASE2_REPORT_GRADE_LOCK_V1,
    status:wmGetCodeGsLmsPhase2ReportGradeLockStatus_
  }),
  PHASE3:Object.freeze({
    registry:WM_CODEGS_LMS_PHASE3_PROFILE_SEND_LOCK_V1,
    status:wmGetCodeGsLmsPhase3ProfileSendLockStatus_
  }),
  PHASE4:Object.freeze({
    registry:WM_CODEGS_LMS_PHASE4_COMMON_ROLE_SPEED_LOCK_V1,
    status:wmGetCodeGsLmsPhase4CommonRoleSpeedLockStatus_
  })
});

function wmGetCodeGsLmsAllLockVerifiedStatus_(){
  var registry=WM_CODEGS_LMS_PHASE5_INTEGRATED_ALL_LOCK_V1;
  var phases={};
  var allVerified=Object.isFrozen(registry);

  ['PHASE1','PHASE2','PHASE3','PHASE4'].forEach(function(key){
    var phase=registry[key]||{};
    var result=null;
    try{
      result=typeof phase.status==='function'?phase.status():null;
    }catch(error){
      result=null;
    }

    var exactRegistry=
      (key==='PHASE1'&&phase.registry===WM_CODEGS_LMS_PHASE1_CORE_LOCK_V1)||
      (key==='PHASE2'&&phase.registry===WM_CODEGS_LMS_PHASE2_REPORT_GRADE_LOCK_V1)||
      (key==='PHASE3'&&phase.registry===WM_CODEGS_LMS_PHASE3_PROFILE_SEND_LOCK_V1)||
      (key==='PHASE4'&&phase.registry===WM_CODEGS_LMS_PHASE4_COMMON_ROLE_SPEED_LOCK_V1);
    var exactStatus=
      (key==='PHASE1'&&phase.status===wmGetCodeGsLmsPhase1CoreLockStatus_)||
      (key==='PHASE2'&&phase.status===wmGetCodeGsLmsPhase2ReportGradeLockStatus_)||
      (key==='PHASE3'&&phase.status===wmGetCodeGsLmsPhase3ProfileSendLockStatus_)||
      (key==='PHASE4'&&phase.status===wmGetCodeGsLmsPhase4CommonRoleSpeedLockStatus_);

    var statusVerified=!!result&&(
      result.allVerified===true||result.verified===true
    );
    var verified=Object.isFrozen(phase)&&
      !!phase.registry&&Object.isFrozen(phase.registry)&&
      exactRegistry&&exactStatus&&statusVerified;

    phases[key]={
      nestedFrozen:Object.isFrozen(phase),
      registryFrozen:!!phase.registry&&Object.isFrozen(phase.registry),
      exactRegistryReference:exactRegistry,
      exactStatusReference:exactStatus,
      statusVerified:statusVerified,
      verified:verified
    };
    allVerified=allVerified&&verified;
  });

  return {
    version:registry.version,
    date:registry.date,
    integratedRegistryFrozen:Object.isFrozen(registry),
    phases:phases,
    allVerified:allVerified
  };
}

/* 🔒 WM CODE.GS LMS PHASE 5 INTEGRATED ALL LOCK VERIFICATION V1 END */
function wmBackfillCurrentProgressStudentInfo() {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('8.현재진행_DB');

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('8.현재진행_DB 데이터 없음');
    return;
  }

  var data = sheet.getDataRange().getDisplayValues();
  var headers = data[0].map(function(h) {
    return String(h || '').trim();
  });

  var idxStudentId = headers.indexOf('학생ID');
  var targets = {
    학생이름: headers.indexOf('학생이름'),
    학교: headers.indexOf('학교'),
    학년: headers.indexOf('학년'),
    Class: headers.indexOf('Class'),
    교사명: headers.indexOf('교사명')
  };

  if (idxStudentId < 0) {
    Logger.log('❌ 학생ID 헤더 없음');
    return;
  }

  var updatedRows = 0;
  var updatedCells = 0;

  for (var i = 1; i < data.length; i++) {
    var studentId = String(data[i][idxStudentId] || '').trim();
    if (!studentId) continue;

    var profile = getStudentBasicInfoForMap_(studentId);
    if (!profile) continue;

    var source = {
      학생이름: String(profile.studentName || '').trim(),
      학교: String(profile.school || '').trim(),
      학년: String(profile.grade || '').trim(),
      Class: String(profile.className || '').trim(),
      교사명: String(profile.teacherName || '').trim()
    };

    var rowChanged = false;

    Object.keys(targets).forEach(function(name) {
      var col = targets[name];
      if (col < 0) return;

      var currentValue = String(data[i][col] || '').trim();
      var newValue = source[name];

      if (!currentValue && newValue) {
        sheet.getRange(i + 1, col + 1).setValue(newValue);
        updatedCells++;
        rowChanged = true;

        Logger.log(
          '보강 / 행 ' + (i + 1) +
          ' / ' + studentId +
          ' / ' + name +
          ' = ' + newValue
        );
      }
    });

    if (rowChanged) updatedRows++;
  }

  Logger.log('==============================');
  Logger.log('보강 학생행수: ' + updatedRows);
  Logger.log('보강 셀수: ' + updatedCells);
  Logger.log('완료 - 학습진행값 수정 없음');
}

function wmDiagnoseStudentManagementRead() {
  var ss = getLmsSpreadsheet_();
  var sheet = ss.getSheetByName('1.학생관리_DB');
  if (!sheet) {
    Logger.log('WM_STUDENT_DB_DIAG=' + JSON.stringify({success:false, spreadsheetId:ss.getId(), sheetFound:false}));
    return {success:false, message:'1.학생관리_DB 시트를 찾을 수 없습니다.'};
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var values = lastRow > 0 && lastColumn > 0
    ? sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues()
    : [];
  var headers = values.length ? values[0].map(function(v){ return String(v || '').trim(); }) : [];
  var studentIdHeaderIndex = headers.indexOf('학생ID');
  var rawStudentIds = [];
  for (var r = 1; r < Math.min(values.length, 11); r++) {
    rawStudentIds.push({
      row:r + 1,
      colB:String((values[r] && values[r][1]) || '').trim(),
      headerStudentId:studentIdHeaderIndex >= 0 ? String(values[r][studentIdHeaderIndex] || '').trim() : ''
    });
  }

  var parsedCount = 0;
  for (var rr = 1; rr < values.length; rr++) {
    var hasValue = false;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      if (String(values[rr][c] || '').trim()) { hasValue = true; break; }
    }
    if (hasValue) parsedCount += 1;
  }

  var result = {
    success:true,
    env:WM_ENV,
    configuredSpreadsheetId:WM_LMS_SPREADSHEET_ID,
    actualSpreadsheetId:ss.getId(),
    sheetName:sheet.getName(),
    lastRow:lastRow,
    lastColumn:lastColumn,
    headerCount:headers.length,
    studentIdHeaderIndex:studentIdHeaderIndex,
    firstHeaders:headers.slice(0, 40),
    parsedCount:parsedCount,
    rawStudentIds:rawStudentIds
  };
  Logger.log('WM_STUDENT_DB_DIAG=' + JSON.stringify(result));
  return result;
}


/* WM_HISTORY_REPAIR_ONCE_V1
 * Apps Script 편집기에서 수동 실행합니다. 로그인/학습/저장/트리거에서는 호출하지 않습니다.
 * 학습이 없는 시간에 실행합니다. 먼저 wmPreviewLearningHistoryRepair의 before/after를 확인합니다.
 * wmApplyLearningHistoryRepair는 히스토리레벨/히스토리횟수 두 셀만 복구합니다.
 */
function wmBuildLearningHistoryRepairPlan_(progressValues, recordValues) {
  var ph = (progressValues[0] || []).map(wmNormalizeHeaderKeyForCurrentProgress_);
  var rh = (recordValues[0] || []).map(wmNormalizeHeaderKeyForCurrentProgress_);
  function column(headers, name) {
    var index = headers.indexOf(wmNormalizeHeaderKeyForCurrentProgress_(name));
    if (index < 0) throw new Error('히스토리 복구 필수 열 없음: ' + name);
    return index;
  }
  var pid = column(ph, '학생ID'), pl = column(ph, '현재레벨');
  var hl = column(ph, '히스토리레벨'), hc = column(ph, '히스토리횟수');
  var rid = column(rh, '학생ID'), rs = column(rh, 'Set_ID');
  var rd = column(rh, '학습날짜'), rc = column(rh, '완료상태');
  var rr = rh.indexOf('완료');
  var logsByStudent = {};
  var seenStudents = {};
  recordValues.slice(1).forEach(function(row, index) {
    var sid = String(row[rid] || '').trim().toUpperCase();
    var setId = normalizeLevelPlainSetId_(row[rs]);
    if (!sid || !setId || !isValidLearningMapSetId_(setId)) return;
    if (!logsByStudent[sid]) logsByStudent[sid] = [];
    logsByStudent[sid].push({
      setId:setId, sortTime:getLearningRecordSortTime_(row[rd]), rowIndex:index + 1,
      complete:wmSafeIsCompleteStatusForCurrentProgress_(row[rc]) || (rr >= 0 && /^(?:[1-3](?:회차|회)?)$/.test(String(row[rr] || '').trim()))
    });
  });
  var changes = [];
  var skipped = [];
  progressValues.slice(1).forEach(function(row, index) {
    var sid = String(row[pid] || '').trim().toUpperCase();
    if (!sid) return;
    if (seenStudents[sid]) throw new Error('현재진행_DB 학생ID 중복: ' + sid);
    seenStudents[sid] = true;
    var level = Number(String(row[pl] || '').replace(/[^0-9]/g, ''));
    if (!(level >= 3 && level <= 13)) {
      skipped.push({studentId:sid, row:index + 2, reason:'현재레벨 확인 불가'});
      return;
    }
    var logs = logsByStudent[sid] || [];
    var rebuilt = wmBuildCurrentProgressLevelHistoryCache_(logs.filter(function(log){ return log.complete; }), level, logs);
    var before = [String(row[hl] || ''), String(row[hc] || '')];
    var oldLevels = before[0].split('|');
    var oldCounts = before[1].split('|');
    var aligned = oldLevels.length === oldCounts.length;
    var merged = {'히스토리레벨':rebuilt.히스토리레벨, '히스토리횟수':rebuilt.히스토리횟수};
    var unresolved = false;
    oldLevels.forEach(function(oldLevel, i) {
      var n = Number(oldLevel);
      if (!(n >= 3 && n <= 13) || n === level) return;
      var hasEvidence = logs.some(function(log){ return Number(log.setId.split('-')[0]) === n; });
      if (!hasEvidence && !aligned) { unresolved = true; return; }
      // 정상 대응된 과거 횟수는 보존합니다. 위치가 어긋난 횟수는 학습기록으로만 복원합니다.
      if (aligned) {
        var pair = wmFastUpdateHistoryForCompletedLevel_(merged, n, oldCounts[i]);
        merged.히스토리레벨 = pair.levels;
        merged.히스토리횟수 = pair.counts;
      }
    });
    if (unresolved) {
      skipped.push({studentId:sid, row:index + 2, reason:'과거 횟수 대응 불일치 및 학습기록 부족'});
      return;
    }
    var ml = merged.히스토리레벨.split('|'), mc = merged.히스토리횟수.split('|');
    var pairs = ml.map(function(v, i){ return {level:Number(v), round:wmNormalizeCurrentProgressRoundValue_(mc[i])}; });
    pairs.sort(function(a,b){ return a.level === level ? 1 : b.level === level ? -1 : a.level-b.level; });
    var after = [pairs.map(function(p){ return p.level; }).join('|'), pairs.map(function(p){ return p.round; }).join('|')];
    if (before[0] !== after[0] || before[1] !== after[1]) changes.push({studentId:sid, row:index + 2, before:before, after:after});
  });
  return {columns:[hl + 1, hc + 1], changes:changes, skipped:skipped};
}

function wmPreviewLearningHistoryRepair() {
  return wmRunLearningHistoryRepair_(false);
}

function wmApplyLearningHistoryRepair() {
  return wmRunLearningHistoryRepair_(true);
}

function wmRunLearningHistoryRepair_(apply) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('다른 관리 작업이 실행 중입니다. 잠시 후 다시 실행해 주세요.');
  try {
    var ss = getLmsSpreadsheet_();
    var progressSheet = ss.getSheetByName('8.현재진행_DB');
    var recordSheet = ss.getSheetByName('2.학습기록_DB');
    if (!progressSheet || !recordSheet) throw new Error('현재진행_DB 또는 학습기록_DB 없음');
    var progressValues = progressSheet.getDataRange().getDisplayValues();
    var plan = wmBuildLearningHistoryRepairPlan_(progressValues, recordSheet.getDataRange().getDisplayValues());
    Logger.log(JSON.stringify({apply:!!apply, spreadsheetId:ss.getId(), plan:plan}));
    if (apply) {
      plan.changes.forEach(function(change) {
        var freshRow = progressSheet.getRange(change.row, 1, 1, progressValues[0].length).getDisplayValues()[0];
        if (JSON.stringify(freshRow) !== JSON.stringify(progressValues[change.row - 1])) {
          throw new Error('복구 중 데이터 변경 감지. 실행 로그를 확인하고 미리보기부터 다시 실행하세요: ' + change.studentId);
        }
        if (plan.columns[1] === plan.columns[0] + 1) {
          progressSheet.getRange(change.row, plan.columns[0], 1, 2).setValues([change.after]);
        } else {
          progressSheet.getRange(change.row, plan.columns[0]).setValue(change.after[0]);
          progressSheet.getRange(change.row, plan.columns[1]).setValue(change.after[1]);
        }
      });
      SpreadsheetApp.flush();
      wmClearRuntimeCachesForStudents_(plan.changes.map(function(change){ return change.studentId; }));
    }
    return {applied:!!apply, changedStudents:plan.changes.length, changes:plan.changes, skipped:plan.skipped};
  } finally {
    lock.releaseLock();
  }
}
