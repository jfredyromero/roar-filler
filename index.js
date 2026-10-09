// ===== This function is injected and runs INSIDE the timesheet page =====
async function autofillTimesheet(dryRun) {
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
    const ENTRIES_HOLIDAY = [
      { hours: 8, groupId: '67', group: 'Cross Capability - Offtime', taskId: '67A', taskName: 'Offtime' },
    ];

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

    log('Detected week:');
    DAYS.forEach(d => log(
      `  ${d.date}${d.holiday ? ' [holiday]' : ''}${!d.enabled ? ' [weekend]' : ''} → ${d.currentHours}h`,
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
      return { logs, errors: 1 };
    }
    log('🔑 Token obtained from MSAL cache.', 'ok');

    if (dryRun) {
      log('— Preview: nothing was saved. —', 'warn');
      DAYS.filter(d => d.enabled).forEach(d => {
        if (d.currentHours >= 8) { log(`↷ ${d.date} would be skipped (${d.currentHours}h)`, 'skip'); return; }
        if (d.currentHours > 0)  { log(`⚠️ ${d.date} partial (${d.currentHours}h): would be skipped`, 'warn'); return; }
        const entries = d.holiday ? ENTRIES_HOLIDAY : ENTRIES_NORMAL;
        entries.forEach(e => log(`→ ${d.date}: ${e.hours}h | ${e.group} | ${e.taskName}`));
      });
      return { logs, errors: 0 };
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

      const job     = day.holiday ? JOB_HOLIDAY     : JOB_NORMAL;
      const entries = day.holiday ? ENTRIES_HOLIDAY : ENTRIES_NORMAL;

      for (const entry of entries) {
        try {
          const r = await saveEntry(day, job, entry);
          if (r.ok) {
            log(`✔ ${day.date} → ${entry.hours}h | ${entry.group}`, 'ok');
          } else {
            errors++;
            log(`✖ ${day.date} → HTTP ${r.status} ${JSON.stringify(r.data || {})}`, 'err');
            if (r.status === 401) { log('Token expired. Reload the page.', 'err'); return { logs, errors }; }
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
    return { logs, errors };
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

  const run = async dryRun => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.url?.includes('roar.marcel.ai')) {
      render([{ msg: 'Open the timesheet page first (roar.marcel.ai).', type: 'err' }]);
      return;
    }
    document.querySelectorAll('button').forEach(b => b.disabled = true);
    render([{ msg: dryRun ? 'Analyzing week…' : 'Saving entries…', type: 'warn' }]);
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        world: 'MAIN',
        func: autofillTimesheet,
        args: [dryRun],
      });
      const allResults = results.map(r => r.result).filter(Boolean);
      const best = allResults.find(r => r.logs.some(l => l.msg.includes('Detected week')));
      const all = best ? best.logs : allResults.flatMap((r, i) => [
        { msg: `--- frame ${i} ---`, type: 'warn' },
        ...r.logs,
      ]);
      render(all);
    } catch (e) {
      render([{ msg: 'Error injecting script: ' + e.message, type: 'err' }]);
    } finally {
      document.querySelectorAll('button').forEach(b => b.disabled = false);
    }
  };

  document.getElementById('preview').addEventListener('click', () => run(true));
  document.getElementById('run').addEventListener('click', () => run(false));
  document.getElementById('reload').addEventListener('click', async () => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    chrome.tabs.reload(tab.id);
    window.close();
  });
