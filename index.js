// ===== This function is injected and runs INSIDE the timesheet page =====
// mode: 'detect' (only read the week) | 'preview' (dry run) | 'run' (save).
// meetingDates: YYYY-MM-DD days that get 2h of meetings instead of 1h.
async function autofillTimesheet(mode, meetingDates = []) {
    const logs = [];
    const log = (msg, type = '') => { logs.push({ msg, type }); console.log(msg); };

    const API = 'https://roar-api.marcel.ai/api/timesheet/api/timesheet/save';
    const EMPLOYEE = '10097729';

    const JOB_NORMAL = {
      jobId: 'H1681-000931-00',
      workCode: 'L01717',
      billableIndicator: 'N',
      task: 'PG01',
      ppmTaskId: '69e794f200008e0f7a14964ac91f753e',
      ppmTaskName: 'Digitas AI: Platform 2026',
    };
    const JOB_HOLIDAY = {
      jobId: 'H7131-000142-04',
      workCode: 'L01717',
      billableIndicator: 'N',
      task: 'PG04',
      ppmTaskId: '69e263990015c0ad70044a1dde380851',
      ppmTaskName: 'Holiday - US8501 - Digitas Agencies',
    };
    const ENTRIES_NORMAL = [
      { hours: 1, groupId: '66', group: 'Cross Capability - General', taskId: '66D', taskName: 'Project Meeting' },
      { hours: 7, groupId: '58', group: 'Tech - Engineering',         taskId: '58A', taskName: 'Development & Engineering Collaboration' },
    ];
    const ENTRIES_MEETING = [
      { ...ENTRIES_NORMAL[0], hours: 2 },
      { ...ENTRIES_NORMAL[1], hours: 6 },
    ];
    const ENTRIES_HOLIDAY = [
      { hours: 8, groupId: '67', group: 'Cross Capability - Offtime', taskId: '67A', taskName: 'Offtime' },
    ];
    const entriesFor = day => day.holiday ? ENTRIES_HOLIDAY
      : meetingDates.includes(day.date) ? ENTRIES_MEETING
      : ENTRIES_NORMAL;

    const pause = ms => new Promise(r => setTimeout(r, ms));

    // ---- Week read from the DOM ----
    const rowByCode = code => {
      const span = [...document.querySelectorAll('.grid-row span')]
        .find(s => s.textContent.trim() === code);
      if (!span) throw new Error('Could not find the row for code ' + code);
      return span.closest('.grid-row');
    };

    // userJobId is specific to each week (rOar creates a new assignment every time
    // you add the job), so it needs to be read live from React props instead of
    // hardcoding it — otherwise saving fails with "Timesheet week should be same
    // with the jobs week".
    const userJobIdFromRow = row => {
      const key = Object.keys(row).find(k => k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$'));
      let node = key ? row[key] : null;
      for (let i = 0; node && i < 25; i++, node = node.return) {
        if (node.memoizedProps?.timesheetJob?.id) return node.memoizedProps.timesheetJob.id;
      }
      return null;
    };

    let DAYS;
    try {
      const normalRow = rowByCode(JOB_NORMAL.jobId);
      JOB_NORMAL.userJobId = userJobIdFromRow(normalRow);
      if (!JOB_NORMAL.userJobId) {
        throw new Error('Could not read this week\'s userJobId from React (did rOar\'s structure change?)');
      }

      const normalInputs = [...normalRow.querySelectorAll('input[type="text"]')]
        .filter(i => /^\d{4}-\d{2}-\d{2}-/.test(i.id));
      // Header = first grid row with no hour inputs that isn't the totals/submit footer
      // (don't anchor to <main>/<nav>: rOar has dropped those wrappers before).
      const headerRow = [...document.querySelectorAll('.grid-row')]
        .find(r => !r.closest('footer') && !r.querySelector('input[type="text"]'));
      const dayCells = [...(headerRow?.querySelectorAll('.grid-row__cell') || [])].slice(1, 8);

      // The Holiday job is only needed that week if there's an actual holiday —
      // otherwise the user never adds it and this shouldn't require that row.
      const hasHoliday = dayCells.some(c => /Public Holiday/i.test(c?.textContent || ''));
      let holidayRow = null;
      if (hasHoliday) {
        holidayRow = rowByCode(JOB_HOLIDAY.jobId);
        JOB_HOLIDAY.userJobId = userJobIdFromRow(holidayRow);
        if (!JOB_HOLIDAY.userJobId) {
          throw new Error('Could not read the Holiday job\'s userJobId from React (did rOar\'s structure change?)');
        }
      }

      DAYS = normalInputs.map((input, i) => {
        const date = input.id.slice(0, 10);
        const holidayInput = holidayRow?.querySelector(`input[id^="${date}-"]`);
        return {
          date,
          dayIndex: i,
          holiday: /Public Holiday/i.test(dayCells[i]?.textContent || ''),
          enabled: !input.disabled,
          currentHours: (parseFloat(input.value) || 0) + (parseFloat(holidayInput?.value) || 0),
        };
      });
    } catch (e) {
      log('⚠️ ' + e.message + ' — are you on the timesheet with the corresponding job added?', 'err');
      return { logs, errors: 1 };
    }

    // `days` marks the frame that actually holds the grid (the popup picks it by this)
    const result = errors => ({ logs, errors, days: DAYS });
    if (mode === 'detect') return result(0);

    log('Detected week:');
    DAYS.forEach(d => log(
      `  ${d.date}${d.holiday ? ' [holiday]' : ''}${!d.enabled ? ' [weekend]' : ''}`
        + `${!d.holiday && meetingDates.includes(d.date) ? ' [2h meeting]' : ''} → ${d.currentHours}h`,
      'skip'
    ));

    // ---- Token from MSAL cache ----
    const getToken = () => {
      const now = Math.floor(Date.now() / 1000);
      const candidates = [];
      for (const storage of [localStorage, sessionStorage]) {
        for (let i = 0; i < storage.length; i++) {
          const key = storage.key(i);
          if (!/accesstoken/i.test(key)) continue;
          try {
            const val = JSON.parse(storage.getItem(key));
            if (val && val.credentialType === 'AccessToken' && val.secret) candidates.push(val);
          } catch (_) {}
        }
      }
      const validTokens = candidates
        .filter(t => (t.target || '').includes('timesheet_api_access'))
        .filter(t => !t.expiresOn || parseInt(t.expiresOn, 10) > now + 60)
        .sort((a, b) => parseInt(b.expiresOn || 0, 10) - parseInt(a.expiresOn || 0, 10));
      return validTokens.length ? validTokens[0].secret : null;
    };

    const TOKEN = getToken();
    if (!TOKEN) {
      log('⚠️ No valid token found. Reload the page and try again.', 'err');
      return result(1);
    }
    log('🔑 Token obtained from MSAL cache.', 'ok');

    if (mode === 'preview') {
      log('— Preview: nothing was saved. —', 'warn');
      DAYS.filter(d => d.enabled).forEach(d => {
        if (d.currentHours >= 8) { log(`↷ ${d.date} would be skipped (${d.currentHours}h)`, 'skip'); return; }
        if (d.currentHours > 0)  { log(`⚠️ ${d.date} partial (${d.currentHours}h): would be skipped`, 'warn'); return; }
        entriesFor(d).forEach(e => log(`→ ${d.date}: ${e.hours}h | ${e.group} | ${e.taskName}`));
      });
      return result(0);
    }

    // ---- Saving ----
    const saveEntry = async (day, job, entry) => {
      const body = {
        jobId: job.jobId, userJobId: job.userJobId, comment: '',
        billableIndicator: job.billableIndicator, dayIndex: day.dayIndex,
        durationValue: entry.hours, employeeCode: EMPLOYEE, workCode: job.workCode,
        reportedDate: day.date, culture: 'en-us', countryCode: '', location: '',
        locationCode: '', id: '', task: job.task, percentageEntry: false,
        ppmTaskId: job.ppmTaskId, ppmTaskName: job.ppmTaskName,
        durationPercentageValue: 0, activity: '', lineManagerCode: null,
        purposeOfOffsiteRTOCompliance: null,
        operationalTaskGroupId: entry.groupId, operationalTaskGroupName: entry.group,
        operationalTaskId: entry.taskId, operationalTaskName: entry.taskName,
      };
      const res = await fetch(API, {
        method: 'POST',
        headers: {
          'accept': 'application/json, text/plain, */*',
          'authorization': 'Bearer ' + TOKEN,
          'content-type': 'application/json',
          'x-app-name': 'timesheet',
        },
        credentials: 'omit',
        body: JSON.stringify(body),
      });
      let data = null;
      try { data = await res.json(); } catch (_) {}
      return { ok: res.ok, status: res.status, data };
    };

    let errors = 0;
    for (const day of DAYS) {
      if (!day.enabled) continue;
      if (day.currentHours >= 8) { log(`↷ ${day.date} already has ${day.currentHours}h`, 'skip'); continue; }
      if (day.currentHours > 0) {
        log(`⚠️ ${day.date} has ${day.currentHours}h (partial): skipping to avoid duplicating`, 'warn');
        continue;
      }

      const job = day.holiday ? JOB_HOLIDAY : JOB_NORMAL;

      for (const entry of entriesFor(day)) {
        try {
          const r = await saveEntry(day, job, entry);
          if (r.ok) {
            log(`✔ ${day.date} → ${entry.hours}h | ${entry.group}`, 'ok');
          } else {
            errors++;
            log(`✖ ${day.date} → HTTP ${r.status} ${JSON.stringify(r.data || {})}`, 'err');
            if (r.status === 401) { log('Token expired. Reload the page.', 'err'); return result(errors); }
          }
        } catch (err) {
          errors++;
          log(`✖ ${day.date} → network error: ${err.message}`, 'err');
        }
        await pause(600);
      }
    }

    log(errors === 0
      ? '✔ Done. Reload the page and verify before submitting.'
      : `⚠ Finished with ${errors} error(s).`, errors === 0 ? 'ok' : 'warn');
    return result(errors);
  }

  // ===== Popup UI =====
  const $log = document.getElementById('log');
  const render = entries => {
    $log.innerHTML = '';
    entries.forEach(({ msg, type }) => {
      const div = document.createElement('div');
      div.textContent = msg;
      if (type) div.className = type;
      $log.appendChild(div);
    });
    $log.scrollTop = $log.scrollHeight;
  };

  // Days picked as "2h meeting", persisted per week so they survive the popup closing
  const $days = document.getElementById('days');
  let meetings = new Set();
  let storageKey = null;

  const saveMeetings = () => {
    if (storageKey) chrome.storage?.session?.set({ [storageKey]: [...meetings] }).catch(() => {});
  };

  const renderDays = days => {
    $days.innerHTML = '';
    days.filter(d => d.enabled).forEach(d => {
      const chip = document.createElement('button');
      chip.textContent = new Date(d.date + 'T00:00')
        .toLocaleDateString('en-US', { weekday: 'short', day: '2-digit' });
      chip.className = 'day' + (meetings.has(d.date) ? ' on' : '');
      chip.disabled = d.holiday || d.currentHours > 0;
      chip.title = d.holiday ? 'Holiday' : d.currentHours > 0 ? `Already has ${d.currentHours}h` : 'Toggle 2h meeting';
      chip.addEventListener('click', () => {
        meetings.has(d.date) ? meetings.delete(d.date) : meetings.add(d.date);
        saveMeetings();
        renderDays(days);
      });
      $days.appendChild(chip);
    });
  };

  const getRoarTab = async () => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab?.url?.includes('roar.marcel.ai')) return tab;
    render([{ msg: 'Open the timesheet page first (roar.marcel.ai).', type: 'err' }]);
    return null;
  };

  // Returns the result of the frame that holds the grid, or every frame's logs if none does
  const inject = async (tab, mode) => {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      world: 'MAIN',
      func: autofillTimesheet,
      args: [mode, [...meetings]],
    });
    const allResults = results.map(r => r.result).filter(Boolean);
    return allResults.find(r => r.days) || {
      logs: allResults.flatMap((r, i) => [{ msg: `--- frame ${i} ---`, type: 'warn' }, ...r.logs]),
    };
  };

  const detectWeek = async () => {
    const tab = await getRoarTab();
    if (!tab) return;
    try {
      const { days, logs } = await inject(tab, 'detect');
      if (!days) { render(logs); return; }
      storageKey = 'meetings:' + days[0].date;
      // Remembering the selection is a nice-to-have: never let storage block the chips
      const stored = await chrome.storage?.session?.get(storageKey).catch(() => null);
      meetings = new Set(stored?.[storageKey] || []);
      renderDays(days);
    } catch (e) {
      render([{ msg: 'Error detecting the week: ' + e.message, type: 'err' }]);
    }
  };

  const run = async mode => {
    const tab = await getRoarTab();
    if (!tab) return;
    const $actions = document.querySelectorAll('#actions button');
    $actions.forEach(b => b.disabled = true);
    render([{ msg: mode === 'preview' ? 'Analyzing week…' : 'Saving entries…', type: 'warn' }]);
    try {
      render((await inject(tab, mode)).logs);
    } catch (e) {
      render([{ msg: 'Error injecting script: ' + e.message, type: 'err' }]);
    } finally {
      $actions.forEach(b => b.disabled = false);
    }
  };

  detectWeek();
  document.getElementById('preview').addEventListener('click', () => run('preview'));
  document.getElementById('run').addEventListener('click', () => run('run'));
  document.getElementById('reload').addEventListener('click', async () => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    chrome.tabs.reload(tab.id);
    window.close();
  });
