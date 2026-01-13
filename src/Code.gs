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
    .addItem('180 days', '180', false);

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

  inputSection.addWidget(phraseInput);
  inputSection.addWidget(scopeInput);
  inputSection.addWidget(ageInput);
  inputSection.addWidget(modeInput);
  inputSection.addWidget(CardService.newKeyValue().setTopLabel('Preview only').setContent('Dry-run mode').setSwitch(dryRunSwitch));
  inputSection.addWidget(CardService.newTextParagraph().setText('Exclusions always on: -is:starred -is:important'));

  var buttonSet = CardService.newButtonSet();

  var previewAction = CardService.newAction().setFunctionName('handlePreview');
  var previewButton = CardService.newTextButton()
    .setText('Preview')
    .setOnClickAction(previewAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  var cleanAction = CardService.newAction().setFunctionName('handleConfirm');
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

  var listResp = safeList_(query, PREVIEW_LIMIT);
  var totalEstimate = listResp.total;
  var ids = listResp.ids;

  var items = ids.map(function(id) {
    return safeGetMetadata_(id);
  });

  return buildPreviewCard_(query, totalEstimate, items, mode, dryRun);
}

function handleConfirm(e) {
  var queryBuild = getQueryFromEvent_(e);
  if (queryBuild.error) {
    return buildValidationResponse_(queryBuild.error);
  }

  var query = queryBuild.query;
  var mode = queryBuild.mode;
  var dryRun = queryBuild.dryRun;

  var listResp = safeList_(query, PREVIEW_LIMIT);
  var totalEstimate = listResp.total;
  var capped = totalEstimate > MAX_PROCESS;
  var effectiveCount = Math.min(totalEstimate, MAX_PROCESS);

  return buildConfirmCard_(query, effectiveCount, capped, mode, dryRun);
}

function handleExecute(e) {
  var query = e.parameters && e.parameters.query ? e.parameters.query : '';
  var mode = e.parameters && e.parameters.mode ? e.parameters.mode : DEFAULTS.MODE;
  var dryRun = isTrue_(e.parameters && e.parameters.dryRun);

  if (!query) {
    return buildValidationResponse_('Missing query. Return to the home card and try again.');
  }

  if (dryRun) {
    return buildResultCard_({
      query: query,
      mode: mode,
      dryRun: true,
      processed: 0,
      success: 0,
      failed: 0,
      capped: false,
      errors: []
    });
  }

  var listResp = listMessageIds_(query, MAX_PROCESS);
  var ids = listResp.ids;
  var capped = listResp.capped;

  var result = processMessages_(ids, mode);
  result.query = query;
  result.mode = mode;
  result.dryRun = false;
  result.capped = capped;

  return buildResultCard_(result);
}

function buildPreviewCard_(query, totalEstimate, items, mode, dryRun) {
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
      section.addWidget(CardService.newTextParagraph().setText(item));
    });
  }

  var buttonSet = CardService.newButtonSet();
  var confirmAction = CardService.newAction()
    .setFunctionName('handleConfirm')
    .setParameters({
      query: query,
      mode: mode,
      dryRun: dryRun ? 'true' : 'false'
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

function buildConfirmCard_(query, count, capped, mode, dryRun) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Confirm'));

  var section = CardService.newCardSection();
  var actionLabel = mode === 'ARCHIVE' ? 'archive' : 'trash';

  if (dryRun) {
    section.addWidget(CardService.newTextParagraph().setText('Dry-run is enabled. No messages will be modified.'));
  }

  section.addWidget(CardService.newTextParagraph().setText('You are about to ' + actionLabel + ' ' + count + ' messages.'));
  section.addWidget(CardService.newTextParagraph().setText('Query: ' + query));

  if (capped) {
    section.addWidget(CardService.newTextParagraph().setText('Only the first ' + MAX_PROCESS + ' messages will be processed. Run again to continue.'));
  }

  var confirmAction = CardService.newAction()
    .setFunctionName('handleExecute')
    .setParameters({
      query: query,
      mode: mode,
      dryRun: dryRun ? 'true' : 'false'
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
  section.addWidget(CardService.newTextParagraph().setText('Query: ' + result.query));
  section.addWidget(CardService.newTextParagraph().setText('Processed: ' + result.processed));
  section.addWidget(CardService.newTextParagraph().setText('Successfully ' + actionLabel + ': ' + result.success));
  section.addWidget(CardService.newTextParagraph().setText('Failed: ' + result.failed));

  if (result.capped) {
    section.addWidget(CardService.newTextParagraph().setText('Only the first ' + MAX_PROCESS + ' messages were processed. Run again to continue.'));
  }

  if (result.errors && result.errors.length) {
    section.addWidget(CardService.newTextParagraph().setText('Errors: ' + result.errors.join(' | ')));
  }

  card.addSection(section);
  return card.build();
}

function getQueryFromEvent_(e) {
  var params = e.parameters || {};
  var form = e.formInput || {};

  if (params.query) {
    return {
      query: params.query,
      mode: params.mode || DEFAULTS.MODE,
      dryRun: isTrue_(params.dryRun)
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
      dryRun: isTrue_(params.dryRun)
    };
  }

  var phrase = (form.phrase || '').trim();
  if (!phrase) {
    return { error: 'Phrase is required.' };
  }

  var scope = form.scope || DEFAULTS.SCOPE;
  var age = form.age || DEFAULTS.AGE_DAYS;
  var mode = form.mode || DEFAULTS.MODE;
  var dryRun = form.dryRun ? true : DEFAULTS.DRY_RUN;

  return {
    query: buildQuery_(phrase, scope, age),
    mode: mode,
    dryRun: dryRun
  };
}

function buildQuery_(phrase, scope, age) {
  var safePhrase = escapeQuery_(phrase);
  var core = scope === 'SUBJECT' ? 'subject:"' + safePhrase + '"' : '"' + safePhrase + '"';
  return core + ' older_than:' + age + 'd ' + EXCLUSIONS;
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

    var from = getHeader_(headers, 'From') || 'Unknown sender';
    var subject = getHeader_(headers, 'Subject') || '(no subject)';
    var date = getHeader_(headers, 'Date') || '';

    return 'From: ' + escapeHtml_(from) + '<br>Subject: ' + escapeHtml_(subject) + '<br>Date: ' + escapeHtml_(date);
  } catch (err) {
    return 'Unable to load message metadata.';
  }
}

function getHeader_(headers, name) {
  var match = headers.filter(function(h) {
    return h.name && h.name.toLowerCase() === name.toLowerCase();
  })[0];
  return match ? match.value : '';
}

function listMessageIds_(query, maxCount) {
  var ids = [];
  var pageToken = null;
  var capped = false;

  do {
    var resp = Gmail.Users.Messages.list('me', {
      q: query,
      maxResults: LIST_PAGE_SIZE,
      pageToken: pageToken
    });

    if (resp && resp.messages) {
      resp.messages.forEach(function(m) {
        if (ids.length < maxCount) {
          ids.push(m.id);
        }
      });
    }

    pageToken = resp.nextPageToken;
    if (ids.length >= maxCount) {
      capped = true;
      break;
    }
  } while (pageToken);

  return { ids: ids, capped: capped };
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
      failed += batch.length;
      if (errors.length < 3) {
        errors.push(shortError_(err));
      }
      if (mode !== 'ARCHIVE') {
        var fallback = trashIndividually_(batch);
        success += fallback.success;
        failed -= fallback.success;
        if (fallback.error && errors.length < 3) {
          errors.push(fallback.error);
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
  var error = '';
  try {
    batch.forEach(function(id) {
      try {
        Gmail.Users.Messages.trash('me', id);
        success += 1;
      } catch (err) {
        if (!error) {
          error = shortError_(err);
        }
      }
    });
  } catch (err) {
    if (!error) {
      error = shortError_(err);
    }
  }
  return { success: success, error: error };
}

function shortError_(err) {
  var message = err && err.message ? err.message : String(err);
  return message.substring(0, 120);
}

function isTrue_(value) {
  return String(value).toLowerCase() === 'true';
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
