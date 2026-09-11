// ===== Esta función se inyecta y se ejecuta DENTRO de la página del timesheet =====
async function autofillTimesheet(dryRun) {
    const logs = [];
    const log = (msg, tipo = '') => { logs.push({ msg, tipo }); console.log(msg); };

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
    const JOB_FESTIVO = {
      jobId: 'H7131-000142-04',
      workCode: 'L01717',
      billableIndicator: 'N',
      task: 'PG04',
      ppmTaskId: '69e263990015c0ad70044a1dde380851',
      ppmTaskName: 'Holiday - US8501 - Digitas Agencies',
    };
    const ENTRIES_NORMAL = [
      { horas: 1, grupoId: '66', grupo: 'Cross Capability - General', tareaId: '66D', tarea: 'Project Meeting' },
      { horas: 7, grupoId: '58', grupo: 'Tech - Engineering',         tareaId: '58A', tarea: 'Development & Engineering Collaboration' },
    ];
    const ENTRIES_FESTIVO = [
      { horas: 8, grupoId: '67', grupo: 'Cross Capability - Offtime', tareaId: '67A', tarea: 'Offtime' },
    ];

    const pausa = ms => new Promise(r => setTimeout(r, ms));

    // ---- Semana leída del DOM ----
    const filaPorCodigo = codigo => {
      const span = [...document.querySelectorAll('main .grid-row span')]
        .find(s => s.textContent.trim() === codigo);
      if (!span) throw new Error('No encontré la fila del código ' + codigo);
      return span.closest('.grid-row');
    };

    // El userJobId es propio de cada semana (rOar crea una asignación nueva cada vez
    // que agregás el job), así que hay que leerlo en vivo desde las props de React
    // en vez de hardcodearlo — si no, el guardado falla con "Timesheet week should
    // be same with the jobs week".
    const userJobIdDeFila = fila => {
      const key = Object.keys(fila).find(k => k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$'));
      let nodo = key ? fila[key] : null;
      for (let i = 0; nodo && i < 25; i++, nodo = nodo.return) {
        if (nodo.memoizedProps?.timesheetJob?.id) return nodo.memoizedProps.timesheetJob.id;
      }
      return null;
    };

    let DIAS;
    try {
      const filaNormal = filaPorCodigo(JOB_NORMAL.jobId);
      JOB_NORMAL.userJobId = userJobIdDeFila(filaNormal);
      if (!JOB_NORMAL.userJobId) {
        throw new Error('No pude leer el userJobId de esta semana desde React (¿cambió la estructura de rOar?)');
      }

      const inputsNormal = [...filaNormal.querySelectorAll('input[type="text"]')]
        .filter(i => /^\d{4}-\d{2}-\d{2}-/.test(i.id));
      const celdasDias = [...document.querySelectorAll('nav .grid-row .grid-row__cell')].slice(1, 8);

      // El job Festivo solo hace falta esa semana si hay algún día feriado real —
      // si no, el usuario nunca lo agrega y no hay que exigir esa fila.
      const hayFestivo = celdasDias.some(c => /Public Holiday/i.test(c?.textContent || ''));
      let filaFestivo = null;
      if (hayFestivo) {
        filaFestivo = filaPorCodigo(JOB_FESTIVO.jobId);
        JOB_FESTIVO.userJobId = userJobIdDeFila(filaFestivo);
        if (!JOB_FESTIVO.userJobId) {
          throw new Error('No pude leer el userJobId del job Festivo desde React (¿cambió la estructura de rOar?)');
        }
      }

      DIAS = inputsNormal.map((input, i) => {
        const fecha = input.id.slice(0, 10);
        const inputFestivo = filaFestivo?.querySelector(`input[id^="${fecha}-"]`);
        return {
          fecha,
          dayIndex: i,
          festivo: /Public Holiday/i.test(celdasDias[i]?.textContent || ''),
          habilitado: !input.disabled,
          horasActuales: (parseFloat(input.value) || 0) + (parseFloat(inputFestivo?.value) || 0),
        };
      });
    } catch (e) {
      log('⚠️ ' + e.message + ' — ¿estás en el timesheet con el job correspondiente agregado?', 'err');
      return { logs, errores: 1 };
    }

    log('Semana detectada:');
    DIAS.forEach(d => log(
      `  ${d.fecha}${d.festivo ? ' [festivo]' : ''}${!d.habilitado ? ' [fin de semana]' : ''} → ${d.horasActuales}h`,
      'skip'
    ));

    // ---- Token desde el cache de MSAL ----
    const obtenerToken = () => {
      const ahora = Math.floor(Date.now() / 1000);
      const candidatos = [];
      for (const storage of [localStorage, sessionStorage]) {
        for (let i = 0; i < storage.length; i++) {
          const key = storage.key(i);
          if (!/accesstoken/i.test(key)) continue;
          try {
            const val = JSON.parse(storage.getItem(key));
            if (val && val.credentialType === 'AccessToken' && val.secret) candidatos.push(val);
          } catch (_) {}
        }
      }
      const validos = candidatos
        .filter(t => (t.target || '').includes('timesheet_api_access'))
        .filter(t => !t.expiresOn || parseInt(t.expiresOn, 10) > ahora + 60)
        .sort((a, b) => parseInt(b.expiresOn || 0, 10) - parseInt(a.expiresOn || 0, 10));
      return validos.length ? validos[0].secret : null;
    };

    const TOKEN = obtenerToken();
    if (!TOKEN) {
      log('⚠️ No hay token vigente. Recarga la página y vuelve a intentar.', 'err');
      return { logs, errores: 1 };
    }
    log('🔑 Token obtenido del cache de MSAL.', 'ok');

    if (dryRun) {
      log('— Vista previa: no se guardó nada. —', 'warn');
      DIAS.filter(d => d.habilitado).forEach(d => {
        if (d.horasActuales >= 8) { log(`↷ ${d.fecha} se saltaría (${d.horasActuales}h)`, 'skip'); return; }
        if (d.horasActuales > 0)  { log(`⚠️ ${d.fecha} parcial (${d.horasActuales}h): se saltaría`, 'warn'); return; }
        const es = d.festivo ? ENTRIES_FESTIVO : ENTRIES_NORMAL;
        es.forEach(e => log(`→ ${d.fecha}: ${e.horas}h | ${e.grupo} | ${e.tarea}`));
      });
      return { logs, errores: 0 };
    }

    // ---- Guardado ----
    const guardarEntry = async (dia, job, entry) => {
      const body = {
        jobId: job.jobId, userJobId: job.userJobId, comment: '',
        billableIndicator: job.billableIndicator, dayIndex: dia.dayIndex,
        durationValue: entry.horas, employeeCode: EMPLOYEE, workCode: job.workCode,
        reportedDate: dia.fecha, culture: 'en-us', countryCode: '', location: '',
        locationCode: '', id: '', task: job.task, percentageEntry: false,
        ppmTaskId: job.ppmTaskId, ppmTaskName: job.ppmTaskName,
        durationPercentageValue: 0, activity: '', lineManagerCode: null,
        purposeOfOffsiteRTOCompliance: null,
        operationalTaskGroupId: entry.grupoId, operationalTaskGroupName: entry.grupo,
        operationalTaskId: entry.tareaId, operationalTaskName: entry.tarea,
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

    let errores = 0;
    for (const dia of DIAS) {
      if (!dia.habilitado) continue;
      if (dia.horasActuales >= 8) { log(`↷ ${dia.fecha} ya tiene ${dia.horasActuales}h`, 'skip'); continue; }
      if (dia.horasActuales > 0) {
        log(`⚠️ ${dia.fecha} tiene ${dia.horasActuales}h (parcial): lo salto para no duplicar`, 'warn');
        continue;
      }

      const job     = dia.festivo ? JOB_FESTIVO     : JOB_NORMAL;
      const entries = dia.festivo ? ENTRIES_FESTIVO : ENTRIES_NORMAL;

      for (const entry of entries) {
        try {
          const r = await guardarEntry(dia, job, entry);
          if (r.ok) {
            log(`✔ ${dia.fecha} → ${entry.horas}h | ${entry.grupo}`, 'ok');
          } else {
            errores++;
            log(`✖ ${dia.fecha} → HTTP ${r.status} ${JSON.stringify(r.data || {})}`, 'err');
            if (r.status === 401) { log('Token expirado. Recarga la página.', 'err'); return { logs, errores }; }
          }
        } catch (err) {
          errores++;
          log(`✖ ${dia.fecha} → error de red: ${err.message}`, 'err');
        }
        await pausa(600);
      }
    }

    log(errores === 0
      ? '✔ Listo. Recarga la página y verifica antes de enviar.'
      : `⚠ Terminado con ${errores} error(es).`, errores === 0 ? 'ok' : 'warn');
    return { logs, errores };
  }

  // ===== UI del popup =====
  const $log = document.getElementById('log');
  const pintar = entradas => {
    $log.innerHTML = '';
    entradas.forEach(({ msg, tipo }) => {
      const div = document.createElement('div');
      div.textContent = msg;
      if (tipo) div.className = tipo;
      $log.appendChild(div);
    });
    $log.scrollTop = $log.scrollHeight;
  };

  const ejecutar = async dryRun => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.url?.includes('roar.marcel.ai')) {
      pintar([{ msg: 'Abre primero la página del timesheet (roar.marcel.ai).', tipo: 'err' }]);
      return;
    }
    document.querySelectorAll('button').forEach(b => b.disabled = true);
    pintar([{ msg: dryRun ? 'Analizando semana…' : 'Guardando entries…', tipo: 'warn' }]);
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        world: 'MAIN',
        func: autofillTimesheet,
        args: [dryRun],
      });
      const resultados = results.map(r => r.result).filter(Boolean);
      const mejor = resultados.find(r => r.logs.some(l => l.msg.includes('Semana detectada')));
      const todos = mejor ? mejor.logs : resultados.flatMap((r, i) => [
        { msg: `--- frame ${i} ---`, tipo: 'warn' },
        ...r.logs,
      ]);
      pintar(todos);
    } catch (e) {
      pintar([{ msg: 'Error al inyectar el script: ' + e.message, tipo: 'err' }]);
    } finally {
      document.querySelectorAll('button').forEach(b => b.disabled = false);
    }
  };

  document.getElementById('preview').addEventListener('click', () => ejecutar(true));
  document.getElementById('run').addEventListener('click', () => ejecutar(false));
  document.getElementById('reload').addEventListener('click', async () => {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    chrome.tabs.reload(tab.id);
    window.close();
  });