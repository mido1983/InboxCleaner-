var DEFAULTS = {
  AGE_DAYS: '14',
  MODE: 'TRASH',
  SCOPE: 'ANY',
  DRY_RUN: true
};

var BATCH_SIZE = 100;
var LIST_PAGE_SIZE = 500;
var PREVIEW_LIMIT = 10;
var MAX_PROCESS = 2000;
var MAX_PER_RUN = 300;
var STATE_TTL_MIN = 30;
var TIME_BUDGET_MS = 25000;
var SLEEP_MS = 300;
var EXCLUSIONS = '-is:starred -is:important';

var PRESETS = {
  promotions: {
    name: 'Promotions',
    query: 'category:promotions older_than:30d'
  },
  social: {
    name: 'Social',
    query: 'category:social older_than:30d'
  },
  noreply: {
    name: 'No-reply',
    query: '("no-reply" OR "do not reply") older_than:30d'
  },
  receipts: {
    name: 'Receipts',
    query: 'subject:(receipt OR invoice) older_than:180d'
  }
};

function onGmailHomepage(e) {
  return buildHomeCard_();
}

function buildHomeCard_() {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('InboxCleaner'));

  var inputSection = CardService.newCardSection().setHeader('Delete what contains');

  var phraseInput = CardService.newTextInput()
    .setFieldName('phrase')
    .setTitle('Phrase')
    .setHint('e.g., unsubscribe');

  var scopeInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Scope')
    .setFieldName('scope')
    .addItem('Anywhere', 'ANY', true)
    .addItem('Subject only', 'SUBJECT', false);

  var ageInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Age')
    .setFieldName('age')
    .addItem('7 days', '7', false)
    .addItem('14 days', '14', true)
    .addItem('30 days', '30', false)
    .addItem('90 days', '90', false)
    .addItem('180 days', '180', false)
    .addItem('All time', 'ALL', false);

  var fromInput = CardService.newTextInput()
    .setFieldName('from')
    .setTitle('From (email or domain)')
    .setHint('e.g., promo@site.com or @site.com');

  var rawQueryInput = CardService.newTextInput()
    .setFieldName('rawQuery')
    .setTitle('Raw Gmail query (advanced)')
    .setHint('e.g., from:@site.com has:attachment before:2024/01/01');

  var modeInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Mode')
    .setFieldName('mode')
    .addItem('Trash', 'TRASH', true)
    .addItem('Archive', 'ARCHIVE', false);

  var dryRunSwitch = CardService.newSwitch()
    .setFieldName('dryRun')
    .setValue('true')
    .setSelected(true)
    .setControlType(CardService.SwitchControlType.SWITCH);

  var autoContinueSwitch = CardService.newSwitch()
    .setFieldName('autoContinue')
    .setValue('true')
    .setSelected(false)
    .setControlType(CardService.SwitchControlType.SWITCH);

  inputSection.addWidget(phraseInput);
  inputSection.addWidget(scopeInput);
  inputSection.addWidget(ageInput);
  inputSection.addWidget(fromInput);
  inputSection.addWidget(rawQueryInput);
  inputSection.addWidget(modeInput);
  inputSection.addWidget(CardService.newKeyValue().setTopLabel('Preview only').setContent('Dry-run mode').setSwitch(dryRunSwitch));
  inputSection.addWidget(CardService.newKeyValue().setTopLabel('Auto-continue').setContent('Process multiple batches per run').setSwitch(autoContinueSwitch));
  inputSection.addWidget(CardService.newTextParagraph().setText('Exclusions always on: -is:starred -is:important'));

  var buttonSet = CardService.newButtonSet();

  var previewAction = CardService.newAction()
    .setFunctionName('handlePreview')
    .setParameters({ dryRun: 'true' });
  var previewButton = CardService.newTextButton()
    .setText('Preview')
    .setOnClickAction(previewAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  var cleanAction = CardService.newAction()
    .setFunctionName('handleConfirm')
    .setParameters({ dryRun: 'false' });
  var cleanButton = CardService.newTextButton()
    .setText('Clean now')
    .setOnClickAction(cleanAction);

  buttonSet.addButton(previewButton);
  buttonSet.addButton(cleanButton);
  inputSection.addWidget(buttonSet);

  card.addSection(inputSection);
  card.addSection(buildPresetsSection_());

  return card.build();
}

function buildPresetsSection_() {
  var section = CardService.newCardSection().setHeader('Presets');

  Object.keys(PRESETS).forEach(function(key) {
    var preset = PRESETS[key];
    var row = CardService.newButtonSet();

    var previewAction = CardService.newAction()
      .setFunctionName('handlePreview')
      .setParameters({ presetId: key, dryRun: 'true', mode: 'TRASH' });
    var previewButton = CardService.newTextButton()
      .setText(preset.name + ' - Preview')
      .setOnClickAction(previewAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

    var cleanAction = CardService.newAction()
      .setFunctionName('handleConfirm')
      .setParameters({ presetId: key, dryRun: 'false', mode: 'TRASH' });
    var cleanButton = CardService.newTextButton()
      .setText(preset.name + ' - Clean now')
      .setOnClickAction(cleanAction);

    row.addButton(previewButton);
    row.addButton(cleanButton);
    section.addWidget(row);
  });

  return section;
}

function handlePreview(e) {
  var queryBuild = getQueryFromEvent_(e);
  if (queryBuild.error) {
    return buildValidationResponse_(queryBuild.error);
  }

  var query = queryBuild.query;
  var mode = queryBuild.mode;
  var dryRun = queryBuild.dryRun;
  var autoContinue = queryBuild.autoContinue;

  var listResp = safeList_(query, PREVIEW_LIMIT);
  var totalEstimate = listResp.total;
  var ids = listResp.ids;

  var items = ids.map(function(id) {
    return safeGetMetadata_(id);
  });

  return buildPreviewCard_(query, totalEstimate, items, mode, dryRun, autoContinue);
}

function handleConfirm(e) {
  var queryBuild = getQueryFromEvent_(e);
  if (queryBuild.error) {
    return buildValidationResponse_(queryBuild.error);
  }

  var query = queryBuild.query;
  var mode = queryBuild.mode;
  var dryRun = queryBuild.dryRun;
  var autoContinue = queryBuild.autoContinue;

  var listResp = safeList_(query, PREVIEW_LIMIT);
  var totalEstimate = listResp.total;
  var capped = totalEstimate > MAX_PROCESS;
  var effectiveCount = Math.min(totalEstimate, MAX_PROCESS);

  return buildConfirmCard_(query, effectiveCount, capped, mode, dryRun, autoContinue);
}

function handleExecute(e) {
  var params = e.parameters || {};
  var stateId = params.stateId || '';
  var state = stateId ? loadRunState_(stateId) : null;
  var query = state && state.query ? state.query : (params.query || '');
  var mode = state && state.mode ? state.mode : (params.mode || DEFAULTS.MODE);
  var dryRun = isTrue_(params.dryRun);
  var pageToken = state && state.pageToken ? state.pageToken : null;
  var processedTotal = state && state.processedTotal ? state.processedTotal : 0;
  var totalEstimate = state && state.totalEstimate ? state.totalEstimate : 0;
  var autoContinue = state && state.autoContinue !== undefined
    ? state.autoContinue
    : isTrue_(params.autoContinue);

  if (!query) {
    return buildValidationResponse_('Missing query. Return to the home card and try again.');
  }

  if (dryRun) {
    return buildResultCard_({
      query: query,
      mode: mode,
      dryRun: true,
      processedThisRun: 0,
      processedTotal: 0,
      totalEstimate: 0,
      success: 0,
      failed: 0,
      capped: false,
      errors: []
    });
  }

  if (!stateId) {
    stateId = Utilities.getUuid();
  }

  var startTime = Date.now();
  var result = { processed: 0, success: 0, failed: 0, errors: [] };
  var nextPageToken = pageToken;
  var capped = false;

  do {
    var listResp = listMessageIds_(query, MAX_PER_RUN, nextPageToken);
    var ids = listResp.ids;
    capped = listResp.capped;
    if (!totalEstimate) {
      totalEstimate = listResp.totalEstimate || 0;
    }

    if (!ids.length) {
      nextPageToken = listResp.nextPageToken;
      break;
    }

    var batchResult = processMessages_(ids, mode);
    result.processed += batchResult.processed;
    result.success += batchResult.success;
    result.failed += batchResult.failed;
    batchResult.errors.forEach(function(err) {
      if (result.errors.length < 3) {
        result.errors.push(err);
      }
    });

    processedTotal += batchResult.processed;
    nextPageToken = listResp.nextPageToken;

    if (!nextPageToken) {
      break;
    }
    if (!autoContinue) {
      break;
    }
  } while (Date.now() - startTime < TIME_BUDGET_MS);

  result.query = query;
  result.mode = mode;
  result.dryRun = false;
  result.capped = capped;
  result.processedThisRun = result.processed;
  result.processedTotal = processedTotal;
  result.totalEstimate = totalEstimate;

  if (nextPageToken) {
    saveRunState_(stateId, {
      query: query,
      mode: mode,
      pageToken: nextPageToken,
      processedTotal: processedTotal,
      totalEstimate: totalEstimate,
      autoContinue: autoContinue
    });
    result.hasMore = true;
    result.stateId = stateId;
  } else {
    clearRunState_(stateId);
    result.hasMore = false;
  }

  return buildResultCard_(result);
}

function buildPreviewCard_(query, totalEstimate, items, mode, dryRun, autoContinue) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Preview'));

  var section = CardService.newCardSection();
  var countText = 'Estimated matches: ' + totalEstimate;
  section.addWidget(CardService.newTextParagraph().setText(countText));
  section.addWidget(CardService.newTextParagraph().setText('Query: ' + query));

  if (totalEstimate > MAX_PROCESS) {
    section.addWidget(CardService.newTextParagraph().setText('Note: Runs are capped at ' + MAX_PROCESS + ' messages per execution.'));
  }

  if (!items.length) {
    section.addWidget(CardService.newTextParagraph().setText('No messages found.'));
  } else {
    items.forEach(function(item) {
      var widget = CardService.newKeyValue()
        .setTopLabel('From: ' + item.from)
        .setContent(item.subject)
        .setBottomLabel('Date: ' + item.date);
      section.addWidget(widget);
    });
  }

  var buttonSet = CardService.newButtonSet();
  var confirmAction = CardService.newAction()
    .setFunctionName('handleConfirm')
    .setParameters({
      query: query,
      mode: mode,
      dryRun: dryRun ? 'true' : 'false',
      autoContinue: autoContinue ? 'true' : 'false'
    });

  var cleanButton = CardService.newTextButton()
    .setText('Clean now')
    .setOnClickAction(confirmAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  buttonSet.addButton(cleanButton);
  section.addWidget(buttonSet);

  card.addSection(section);
  return card.build();
}

function buildConfirmCard_(query, count, capped, mode, dryRun, autoContinue) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Confirm'));

  var section = CardService.newCardSection();
  var actionLabel = mode === 'ARCHIVE' ? 'archive' : 'trash';

  if (dryRun) {
    section.addWidget(CardService.newTextParagraph().setText('Dry-run is enabled. No messages will be modified.'));
  }

  section.addWidget(CardService.newTextParagraph().setText('You are about to ' + actionLabel + ' ' + count + ' messages.'));
  section.addWidget(CardService.newTextParagraph().setText('Query: ' + query));

  if (query.indexOf('older_than:') === -1) {
    section.addWidget(CardService.newTextParagraph().setText('All time can require multiple runs. Use Continue until finished.'));
  }

  if (capped) {
    section.addWidget(CardService.newTextParagraph().setText('Only the first ' + MAX_PROCESS + ' messages will be processed. Run again to continue.'));
  }

  var confirmAction = CardService.newAction()
    .setFunctionName('handleExecute')
    .setParameters({
      query: query,
      mode: mode,
      dryRun: dryRun ? 'true' : 'false',
      autoContinue: autoContinue ? 'true' : 'false'
    });

  var confirmButton = CardService.newTextButton()
    .setText('Confirm and clean')
    .setOnClickAction(confirmAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  section.addWidget(confirmButton);
  card.addSection(section);
  return card.build();
}

function buildResultCard_(result) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Result'));

  var section = CardService.newCardSection();

  if (result.dryRun) {
    section.addWidget(CardService.newTextParagraph().setText('Dry-run is enabled. No messages were modified.'));
  }

  var actionLabel = result.mode === 'ARCHIVE' ? 'archived' : 'trashed';
  var totalEstimate = result.totalEstimate || 0;
  section.addWidget(CardService.newTextParagraph().setText('Query: ' + result.query));
  section.addWidget(CardService.newTextParagraph().setText('Estimated total matches: ' + totalEstimate));
  section.addWidget(CardService.newTextParagraph().setText('Processed this run: ' + result.processedThisRun));
  section.addWidget(CardService.newTextParagraph().setText('Processed total: ' + result.processedTotal));
  section.addWidget(CardService.newTextParagraph().setText('Remaining (est): ' + Math.max(totalEstimate - result.processedTotal, 0)));
  section.addWidget(CardService.newTextParagraph().setText('Successfully ' + actionLabel + ': ' + result.success));
  section.addWidget(CardService.newTextParagraph().setText('Failed: ' + result.failed));

  if (result.capped) {
    section.addWidget(CardService.newTextParagraph().setText('Only the first ' + MAX_PROCESS + ' messages were processed. Run again to continue.'));
  }

  if (result.errors && result.errors.length) {
    section.addWidget(CardService.newTextParagraph().setText('Errors: ' + result.errors.join(' | ')));
  }

  if (result.hasMore && result.stateId) {
    var continueAction = CardService.newAction()
      .setFunctionName('handleExecute')
      .setParameters({ stateId: result.stateId });
    var continueButton = CardService.newTextButton()
      .setText('Continue')
      .setOnClickAction(continueAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
    section.addWidget(continueButton);
  }

  card.addSection(section);
  return card.build();
}

function getQueryFromEvent_(e) {
  var params = e.parameters || {};
  var form = e.formInput || {};
  var paramDryRun = params.dryRun !== undefined ? isTrue_(params.dryRun) : null;
  var paramAutoContinue = params.autoContinue !== undefined ? isTrue_(params.autoContinue) : null;

  if (params.query) {
    return {
      query: params.query,
      mode: params.mode || DEFAULTS.MODE,
      dryRun: paramDryRun !== null ? paramDryRun : false,
      autoContinue: paramAutoContinue !== null ? paramAutoContinue : false
    };
  }

  if (params.presetId) {
    var preset = PRESETS[params.presetId];
    if (!preset) {
      return { error: 'Unknown preset.' };
    }
    return {
      query: preset.query + ' ' + EXCLUSIONS,
      mode: params.mode || DEFAULTS.MODE,
      dryRun: paramDryRun !== null ? paramDryRun : false,
      autoContinue: paramAutoContinue !== null ? paramAutoContinue : false
    };
  }

  var scope = form.scope || DEFAULTS.SCOPE;
  var age = form.age || DEFAULTS.AGE_DAYS;
  var mode = form.mode || DEFAULTS.MODE;
  var dryRun = paramDryRun !== null
    ? paramDryRun
    : (String(form.dryRun).toLowerCase() === 'true');
  var autoContinue = paramAutoContinue !== null
    ? paramAutoContinue
    : (String(form.autoContinue).toLowerCase() === 'true');

  var queryBuild = buildQueryFromInputs_({
    phrase: form.phrase,
    scope: scope,
    age: age,
    from: form.from,
    rawQuery: form.rawQuery
  });
  if (queryBuild.error) {
    return { error: queryBuild.error };
  }

  return {
    query: queryBuild.query,
    mode: mode,
    dryRun: dryRun,
    autoContinue: autoContinue
  };
}

function buildQueryFromInputs_(form) {
  var rawQuery = (form.rawQuery || '').trim();
  if (rawQuery) {
    return { query: rawQuery, isRaw: true };
  }

  var phrase = (form.phrase || '').trim();
  if (!phrase) {
    return { error: 'Phrase is required.' };
  }

  var safePhrase = escapeQuery_(phrase);
  var core = form.scope === 'SUBJECT' ? 'subject:"' + safePhrase + '"' : '"' + safePhrase + '"';
  var fromValue = (form.from || '').trim();
  var fromClause = fromValue ? ' from:' + fromValue : '';
  var ageClause = (!form.age || form.age === 'ALL') ? '' : ' older_than:' + form.age + 'd';
  return { query: core + fromClause + ageClause + ' ' + EXCLUSIONS, isRaw: false };
}

function escapeQuery_(text) {
  return text.replace(/"/g, '\\"');
}

function safeList_(query, maxResults) {
  try {
    var resp = Gmail.Users.Messages.list('me', {
      q: query,
      maxResults: maxResults
    });

    var ids = [];
    if (resp && resp.messages) {
      ids = resp.messages.map(function(m) { return m.id; });
    }

    return {
      ids: ids,
      total: resp ? (resp.resultSizeEstimate || 0) : 0
    };
  } catch (err) {
    return { ids: [], total: 0 };
  }
}

function safeGetMetadata_(id) {
  try {
    var msg = Gmail.Users.Messages.get('me', id, {
      format: 'metadata',
      metadataHeaders: ['From', 'Subject', 'Date']
    });
    var headers = msg.payload && msg.payload.headers ? msg.payload.headers : [];

    var from = escapeHtml_(getHeader_(headers, 'From') || 'Unknown sender');
    var subject = escapeHtml_(getHeader_(headers, 'Subject') || '(no subject)');
    var date = escapeHtml_(getHeader_(headers, 'Date') || '');

    return { from: from, subject: subject, date: date };
  } catch (err) {
    return { from: 'Unknown sender', subject: 'Unable to load message metadata.', date: '' };
  }
}

function getHeader_(headers, name) {
  var match = headers.filter(function(h) {
    return h.name && h.name.toLowerCase() === name.toLowerCase();
  })[0];
  return match ? match.value : '';
}

function listMessageIds_(query, maxCount, pageToken) {
  var ids = [];
  var capped = false;
  var token = pageToken || null;
  var nextPageToken = null;
  var totalEstimate = 0;

  do {
    var resp = Gmail.Users.Messages.list('me', {
      q: query,
      maxResults: LIST_PAGE_SIZE,
      pageToken: token
    });

    if (!totalEstimate && resp && resp.resultSizeEstimate !== undefined) {
      totalEstimate = resp.resultSizeEstimate || 0;
    }

    if (resp && resp.messages) {
      resp.messages.forEach(function(m) {
        if (ids.length < maxCount) {
          ids.push(m.id);
        }
      });
    }

    nextPageToken = resp.nextPageToken;
    if (ids.length >= maxCount) {
      capped = true;
      break;
    }
    token = nextPageToken;
  } while (token);

  return { ids: ids, nextPageToken: nextPageToken, capped: capped, totalEstimate: totalEstimate };
}

function processMessages_(ids, mode) {
  var success = 0;
  var failed = 0;
  var errors = [];

  for (var i = 0; i < ids.length; i += BATCH_SIZE) {
    var batch = ids.slice(i, i + BATCH_SIZE);

    try {
      if (mode === 'ARCHIVE') {
        Gmail.Users.Messages.batchModify({
          ids: batch,
          removeLabelIds: ['INBOX']
        }, 'me');
        success += batch.length;
      } else {
        Gmail.Users.Messages.batchModify({
          ids: batch,
          addLabelIds: ['TRASH']
        }, 'me');
        success += batch.length;
      }
    } catch (err) {
      if (errors.length < 3) {
        errors.push(shortError_(err));
      }
      if (mode === 'ARCHIVE') {
        failed += batch.length;
      } else {
        var trashFallback = trashIndividually_(batch);
        success += trashFallback.success;
        failed += trashFallback.failed;
        if (trashFallback.error && errors.length < 3) {
          errors.push(trashFallback.error);
        }
      }
    }

    Utilities.sleep(SLEEP_MS);
  }

  return {
    processed: ids.length,
    success: success,
    failed: failed,
    errors: errors
  };
}

function trashIndividually_(batch) {
  var success = 0;
  var failed = 0;
  var error = '';
  try {
    batch.forEach(function(id) {
      try {
        Gmail.Users.Messages.trash('me', id);
        success += 1;
      } catch (err) {
        failed += 1;
        if (!error) {
          error = shortError_(err);
        }
      }
    });
  } catch (err) {
    failed += batch.length;
    if (!error) {
      error = shortError_(err);
    }
  }
  return { success: success, failed: failed, error: error };
}

function shortError_(err) {
  var message = err && err.message ? err.message : String(err);
  return message.substring(0, 120);
}

function isTrue_(value) {
  return String(value).toLowerCase() === 'true';
}

function saveRunState_(stateId, state) {
  var props = PropertiesService.getUserProperties();
  var payload = {
    query: state.query,
    mode: state.mode,
    pageToken: state.pageToken,
    processedTotal: state.processedTotal,
    totalEstimate: state.totalEstimate,
    autoContinue: state.autoContinue,
    updatedAt: Date.now()
  };
  props.setProperty('run_' + stateId, JSON.stringify(payload));
}

function loadRunState_(stateId) {
  var props = PropertiesService.getUserProperties();
  var raw = props.getProperty('run_' + stateId);
  if (!raw) {
    return null;
  }
  try {
    var state = JSON.parse(raw);
    if (STATE_TTL_MIN && state.updatedAt) {
      var ageMs = Date.now() - state.updatedAt;
      if (ageMs > STATE_TTL_MIN * 60 * 1000) {
        clearRunState_(stateId);
        return null;
      }
    }
    return state;
  } catch (err) {
    clearRunState_(stateId);
    return null;
  }
}

function clearRunState_(stateId) {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty('run_' + stateId);
}

function buildValidationResponse_(message) {
  var card = buildHomeCard_();
  var notification = CardService.newNotification().setText(message);
  return CardService.newActionResponseBuilder()
    .setNotification(notification)
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function escapeHtml_(text) {
  if (!text) {
    return '';
  }
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
