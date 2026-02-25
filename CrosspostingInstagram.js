const CONFIG = {
  SHEET_NAME: 'Интерфейс',

  DROPDOWN_CELL_FETCH: 'B1',
  FETCH_TRIGGER_VALUE: 'Да',
  FETCH_DEFAULT_VALUE: 'Получить последний пост Instagram',

  DROPDOWN_CELL_PUBLISH: 'A1',
  PUBLISH_TRIGGER_VALUE: 'Да',
  PUBLISH_DEFAULT_VALUE: 'Выложить пост из инстаграм в остальные соцсети?',

  DROPDOWN_CELL_ADAPT: 'D1',
  ADAPT_TRIGGER_VALUE: 'Да',
  ADAPT_DEFAULT_VALUE: 'Адаптировать текст для включенных соцсетей?',

  DROPDOWN_CELL_AUTO: 'C1',
  AUTO_ON_VALUE: 'Автоматическая выкладка ВКЛ',
  AUTO_OFF_VALUE: 'Автоматическая выкладка ВЫКЛ',

  REPOST_DEFAULT_VALUE: 'Перевыложить?',
  REWRITE_DEFAULT_VALUE: 'Переписать текст?',

  OUTPUT_ROW: 3,
  COL_IMAGE: 1,
  COL_CAPTION: 2,
  COL_TYPE: 3,
  COL_MEDIA_URLS: 4,
  COL_INFO: 5,
  COL_POST_ID: 6,

  DROPDOWN_CELL_FETCH_BY_LINK: 'E1',
  FETCH_BY_LINK_TRIGGER_VALUE: 'Да',
  FETCH_BY_LINK_DEFAULT_VALUE: 'Получить пост Instagram по ссылке?',

  INSTAGRAM_LINK_CELL: 'F1',
  SCHEDULE_DATE_CELL: 'G3',
  LINK_CELL: 'E14',

  API_ENDPOINT: 'https://instagram120.p.rapidapi.com/api/instagram/links',
  API_MEDIA_BY_SHORTCODE_ENDPOINT: 'https://instagram120.p.rapidapi.com/api/instagram/mediaByShortcode',

  RAPIDAPI_HOST: 'instagram120.p.rapidapi.com',
  COL_INSTAGRAM_URL: 1,
  ROW_INSTAGRAM_URL: 14,

  SOCIAL_LAYOUT: {
    HEADER_ROW: 4,
    TOGGLE_ROW: 5,
    MODEL_ROW: 6,
    REWRITE_FLAG_ROW: 7,
    REPOST_FLAG_ROW: 8,

    TITLE_ROW: 10,
    PROMPT_ROW: 9,
    OUTPUT_ROW: 11,

    FIRST_COL: 2
  },

  SYSTEM_PROMPT_CELL: 'B14',
  SYSTEM_PROMPT_SUFFIX: 'Всегда отвечай СТРОГО одним JSON-объектом вида {"title": "...", "text": "..."}, ' +
    'без каких-либо комментариев, префиксов, суффиксов и форматирования.',

  RAPIDAPI_QUOTA_CELL: 'C14',
  CLOUDINARY_USAGE_CELL: 'D14',
  // Дата регистрации (гггг-мм-дд) для расчета 30-дневных циклов.
  // Установите корректный год (в симуляции сейчас 2026, значит регистрация была в конце 2025)
  CLOUDINARY_REG_DATE: '2025-12-22'
};

const PROP_KEYS = {
  RAPIDAPI_KEY: 'RAPIDAPI_KEY',
  OPENAI_API_KEY: 'OPENAI_API_KEY',
  CLOUDINARY_CLOUD_NAME: 'CLOUDINARY_CLOUD_NAME',
  CLOUDINARY_API_KEY: 'CLOUDINARY_API_KEY',
  CLOUDINARY_API_SECRET: 'CLOUDINARY_API_SECRET'
};

const ALWAYS_SLIDESHOW_NETWORKS = ['Reddit'];

const SOCIAL_FILTERS = {
  ONLY_REELS: ['Pinterest_reels'],
  EXCLUDE_REELS: ['Pinterest_pictures']
};

const MIXED_SLIDESHOW_NETWORKS = ['TikTok', 'Pinterest_reels', 'Pinterest_pictures', 'LinkedIn'];

const PPM_CONFIG = {
  BASE_URL: 'https://api.postmypost.io/v4.1',
  PROJECT_ID: 320499,
  TOKEN_PROPERTY: 'PPM_ACCESS_TOKEN',

  ACCOUNT_IDS: {
    'Threads': 2051102,
    'Reddit': 2051110,
    'VK': 2050992,
    //'Rutube': *****,
    'Pinterest_pictures': 2051031,
    'Telegram': 2051184,
    'Instagram': 2050988,
    'TikTok': 2051262,
    'LinkedIn': 2051109,
    'YouTube': 2051009,
    'Pinterest_reels': 2051051
  }
};

function autoCheckAndCrosspost() {
  autoCheckAndCrosspost_();
}

function nowString_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd.MM.yyyy HH:mm:ss'
  );
}

function setInfo_(msg) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  sheet.getRange(CONFIG.OUTPUT_ROW, CONFIG.COL_INFO)
    .setValue(msg + ' [' + nowString_() + ']');
}

function jitterInstagramUsernameCase_(url) {
  if (!url) return url;

  var marker = 'instagram.com/';
  var idx = url.indexOf(marker);
  if (idx === -1) {
    return url;
  }

  var prefix = url.substring(0, idx + marker.length);
  var rest = url.substring(idx + marker.length);

  var cut = rest.length;
  var slashIdx = rest.indexOf('/');
  if (slashIdx !== -1 && slashIdx < cut) cut = slashIdx;
  var qIdx = rest.indexOf('?');
  if (qIdx !== -1 && qIdx < cut) cut = qIdx;

  var username = rest.substring(0, cut);
  var suffix = rest.substring(cut);

  if (!username) return url;
  var chars = username.split('');
  for (var i = 0; i < chars.length; i++) {
    var ch = chars[i];
    if (/[a-zA-Z]/.test(ch)) {
      chars[i] = (Math.random() < 0.5)
        ? ch.toLowerCase()
        : ch.toUpperCase();
    }
  }
  var newUsername = chars.join('');

  var newUrl = prefix + newUsername + suffix;
  Logger.log('Instagram URL (jittered): ' + newUrl);
  return newUrl;
}

function getSystemPrompt_(sheet) {
  var userPart = sheet.getRange(CONFIG.SYSTEM_PROMPT_CELL).getValue();
  return (userPart || '') + ' ' + CONFIG.SYSTEM_PROMPT_SUFFIX;
}

function getSocialRowsFromSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');
  }

  var layout = CONFIG.SOCIAL_LAYOUT;

  var lastCol = sheet.getLastColumn();
  var headerRowValues = sheet
    .getRange(layout.HEADER_ROW, layout.FIRST_COL, 1, lastCol - layout.FIRST_COL + 1)
    .getValues()[0];

  var width = headerRowValues.length;
  if (width <= 0) {
    return [];
  }

  var toggleRowValues = sheet
    .getRange(layout.TOGGLE_ROW, layout.FIRST_COL, 1, width)
    .getValues()[0];

  var modelRowValues = sheet
    .getRange(layout.MODEL_ROW, layout.FIRST_COL, 1, width)
    .getValues()[0];

  var titleRowValues = sheet
    .getRange(layout.TITLE_ROW, layout.FIRST_COL, 1, width)
    .getValues()[0];

  var promptRowValues = sheet
    .getRange(layout.PROMPT_ROW, layout.FIRST_COL, 1, width)
    .getValues()[0];

  var outputRowValues = sheet
    .getRange(layout.OUTPUT_ROW, layout.FIRST_COL, 1, width)
    .getValues()[0];

  var socials = [];

  for (var i = 0; i < width; i++) {
    var name = String(headerRowValues[i] || '').trim();
    if (!name) {
      continue;
    }

    var enabledRaw = String(toggleRowValues[i] || '').toLowerCase();
    var enabled = (enabledRaw === 'вкл' || enabledRaw === 'да' || enabledRaw === 'on');

    var model = String(modelRowValues[i] || '').trim();
    var title = String(titleRowValues[i] || '').trim();
    var prompt = String(promptRowValues[i] || '').trim();
    var text = String(outputRowValues[i] || '').trim();

    var col = layout.FIRST_COL + i;

    socials.push({
      name: name,
      col: col,
      enabled: enabled,
      model: model,
      title: title,
      prompt: prompt,
      text: text
    });
  }

  return socials;
}

function buildPostmypostPublicationRequest_(mediaUrls, originalCaption, options) {
  options = options || {};
  var linkToPin = options.link || null;

  var socials = getSocialRowsFromSheet_().filter(function (s) {
    if (options.onlyColumn && s.col === options.onlyColumn) {
      return true;
    }

    if (!s.enabled) return false;

    if (options.onlyColumn && s.col !== options.onlyColumn) return false;
    if (options.onlyName && s.name !== options.onlyName) return false;

    return true;
  });

  if (!socials.length) {
    throw new Error('Нет ни одной соцсети для публикации (проверь Вкл/Выкл и фильтры onlyColumn/onlyName).');
  }

  var hasVideo = false;
  var hasImage = false;
  mediaUrls.forEach(function (u) {
    if (isVideoUrl_(u)) hasVideo = true;
    else hasImage = true;
  });

  var isMixed = hasVideo && hasImage;
  var isCarousel = hasVideo || hasImage;
  var fileIdsCache = {};

  var getFileIds = function (variant) {
    if (fileIdsCache[variant]) return fileIdsCache[variant];

    if (variant === 'original') {
      var ids = uploadMediaToPostmypostByUrls_(mediaUrls);
      fileIdsCache['original'] = ids;
      return ids;
    }

    if (variant === 'slideshow') {
      Logger.log('Генерация слайдшоу через Cloudinary...');
      var slideshowUrl = createCloudinarySlideshowUrl_(mediaUrls);

      if (!slideshowUrl) {
        throw new Error('Cloudinary slideshow URL пустой – не удалось собрать слайдшоу');
      }

      Logger.log('Ссылка на слайдшоу: ' + slideshowUrl);
      var ids = uploadMediaToPostmypostByUrls_([slideshowUrl]);
      fileIdsCache['slideshow'] = ids;
      return ids;
    }


    return [];
  };

  var accountIds = [];
  var details = [];
  var lastError = null;

  socials.forEach(function (s) {
    var accountId = PPM_CONFIG.ACCOUNT_IDS[s.name];
    if (!accountId) {
      Logger.log('Для соцсети "' + s.name + '" не задан account_id в PPM_CONFIG.ACCOUNT_IDS — пропускаю.');
      return;
    }

    // --- ФИЛЬТРАЦИЯ ПО ТИПУ KONТЕНТА (ONLY_REELS / EXCLUDE_REELS) ---
    var isSingleVideo = (mediaUrls.length === 1 && isVideoUrl_(mediaUrls[0]));

    if ((SOCIAL_FILTERS.ONLY_REELS || []).indexOf(s.name) !== -1) {
      if (!isSingleVideo) {
        Logger.log('Пропускаю ' + s.name + ': эта сеть принимает ТОЛЬКО рилсы (single video), а тут другой контент.');
        return;
      }
    }
    if ((SOCIAL_FILTERS.EXCLUDE_REELS || []).indexOf(s.name) !== -1) {
      if (isSingleVideo) {
        Logger.log('Пропускаю ' + s.name + ': эта сеть ИСКЛЮЧАЕТ рилсы (принимает картинки/карусели).');
        return;
      }
    }
    // ----------------------------------------------------------------

    var useSlideshow = false;

    if (ALWAYS_SLIDESHOW_NETWORKS.indexOf(s.name) !== -1) {
      if (mediaUrls.length > 1) {
        useSlideshow = true;
      }
    } else if (MIXED_SLIDESHOW_NETWORKS.indexOf(s.name) !== -1) {
      if (isMixed && mediaUrls.length > 1) {
        useSlideshow = true;
      }
    }
    var currentFileIds = [];
    try {
      currentFileIds = getFileIds(useSlideshow ? 'slideshow' : 'original');
    } catch (e) {
      Logger.log('Ошибка подготовки файлов (' + (useSlideshow ? 'slideshow' : 'original') + ') для ' + s.name + ': ' + e);
      lastError = e;
      return;
    }
    accountIds.push(accountId);

    var contentText = s.text || originalCaption || '';

    var pubType = 1;
    if (s.name === 'YouTube' || s.name === 'Rutube') {
      pubType = 4;
    } else if (s.name === 'VK' && isSingleVideo) {
      pubType = 1;
    }

    var detail = {
      account_id: accountId,
      publication_type: pubType,
      content: contentText,
      file_ids: currentFileIds
    };

    if (s.title) {
      detail.title = s.title;
    }

    if (s.name.toLowerCase() === 'tiktok') {
      detail.tiktok_privacy_status = 1;
      detail.tiktok_comment = true;
      detail.tiktok_duet = true;
      detail.tiktok_stitch = true;
    }

    // Добавляем ссылку к пинам (Pinterest), если она есть
    if (linkToPin && s.name.toLowerCase().indexOf('pinterest') !== -1) {
      detail.link = linkToPin;
    }

    details.push(detail);
  });

  if (!details.length) {
    if (lastError) throw lastError;
    throw new Error(
      'Не удалось подготовить медиа для публикации. ' +
      'Попробуйте ещё раз или проверьте формат исходных файлов.'
    );
  }

  var postAt = options.postAt || Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );

  return {
    project_id: PPM_CONFIG.PROJECT_ID,
    post_at: postAt,
    account_ids: accountIds,
    publication_status: 5,
    details: details
  };
}

function isVideoUrl_(url) {
  if (!url) return false;
  var u = url.toLowerCase().split('?')[0];
  return u.endsWith('.mp4') || u.endsWith('.mov') || u.endsWith('.avi') || u.endsWith('.webm');
}

function getCloudinaryConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    cloudName: props.getProperty(PROP_KEYS.CLOUDINARY_CLOUD_NAME),
    apiKey: props.getProperty(PROP_KEYS.CLOUDINARY_API_KEY),
    apiSecret: props.getProperty(PROP_KEYS.CLOUDINARY_API_SECRET)
  };
}

function cloudinarySign_(params, apiSecret) {
  var keys = Object.keys(params).sort();
  var toSign = keys.map(function (k) {
    return k + '=' + params[k];
  }).join('&') + apiSecret;

  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    toSign
  );

  var sig = bytes.reduce(function (str, b) {
    var v = (b & 0xff).toString(16);
    return str + (v.length === 1 ? '0' + v : v);
  }, '');

  return sig;
}

function uploadToCloudinary_(blob, resourceType) {
  var conf = getCloudinaryConfig_();
  if (!conf.cloudName || !conf.apiKey || !conf.apiSecret) {
    throw new Error('Нет настроек Cloudinary в Script Properties!');
  }

  var url = 'https://api.cloudinary.com/v1_1/' + conf.cloudName + '/' + resourceType + '/upload';

  var timestamp = unixTimestamp_(); // <-- СТРОКА
  var signature = cloudinarySignature_({ timestamp: timestamp }, conf.apiSecret);

  var formData = {
    file: blob,               // можно: blob.setName('upload')
    api_key: String(conf.apiKey),
    timestamp: timestamp,     // <-- СТРОКА
    signature: signature
  };

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: formData,
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var txt = resp.getContentText();

  if (code !== 200 && code !== 201) {
    throw new Error('Cloudinary upload error (' + code + '): ' + txt);
  }

  var data = JSON.parse(txt);
  if (!data.public_id) {
    throw new Error('Cloudinary upload: нет public_id в ответе: ' + txt);
  }

  // Возвращаем объект с метаданными, так как нужны размеры
  return {
    publicId: data.public_id,
    width: data.width,
    height: data.height,
    format: data.format
  };
}

function unixTimestamp_() {
  return String(Math.floor(Date.now() / 1000));
}

function unixTimestampStr_() {
  return String(Math.floor(Date.now() / 1000)); // "1766575232"
}

function cloudinaryStringToSign_(params) {
  // сортировка по ключу как требует Cloudinary
  var keys = Object.keys(params).sort();
  return keys.map(function (k) {
    return k + '=' + String(params[k]);
  }).join('&');
}

function cloudinarySignature_(params, apiSecret) {
  // Cloudinary: параметры сортируются по ключу и конкатенируются как key=value&...
  var keys = Object.keys(params).sort();
  var base = keys.map(function (k) {
    return k + '=' + String(params[k]);
  }).join('&');

  return sha1Hex_(base + apiSecret);
}

function sha1HexUtf8_(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    str,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    if (b < 0) b += 256;
    var h = b.toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}

function formUrlEncode_(params) {
  return Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k]));
  }).join('&');
}

function createCloudinarySlideshowUrl_(urls) {
  var conf = getCloudinaryConfig_();
  if (!conf.cloudName || !conf.apiKey || !conf.apiSecret) {
    throw new Error('Нет настроек Cloudinary в Script Properties!');
  }
  if (!urls || !urls.length) throw new Error('Нет URL для слайдшоу');

  Logger.log('Генерация слайдшоу (v5 - overlay + adaptive AR)...');

  // 1) Загружаем все исходники
  var rawAssets = urls.map(function (u, index) {
    if (!u) return null;
    var isVid = isVideoUrl_(u);
    var media = downloadMediaFromUrl_(u);
    // uploadToCloudinary_ теперь возвращает объект {publicId, width, height...}
    var res = uploadToCloudinary_(media.blob, isVid ? 'video' : 'image');
    Logger.log('Loaded [' + index + '] ' + (isVid ? 'video' : 'image') + ': ' + res.publicId + ' (' + res.width + 'x' + res.height + ')');
    return {
      type: isVid ? 'video' : 'image',
      publicId: res.publicId,
      width: res.width,
      height: res.height,
      index: index
    };
  }).filter(Boolean);

  if (!rawAssets.length) throw new Error('Нет медиа');

  // Определяем базовые размеры по первому ВИДЕО (или, если нет видео, по первому ассету)
  // В смешанном режиме первое видео задает тон.
  var baseVideo = rawAssets.filter(function (a) { return a.type === 'video'; })[0] || rawAssets[0];
  var TARGET_W = baseVideo.width;
  var TARGET_H = baseVideo.height;
  var donorVideoId = baseVideo.type === 'video' ? baseVideo.publicId : 'sample';

  Logger.log('Целевые размеры слайдшоу: ' + TARGET_W + 'x' + TARGET_H + ' (по ассету ' + baseVideo.publicId + ')');

  // 2) Конвертируем все IMAGE в VIDEO-сегменты (3.5 сек)
  // Используем целевые размеры для подложки
  var videoAssets = rawAssets.map(function (asset) {
    if (asset.type === 'video') {
      return asset;
    } else {
      // Это картинка. Создаём сегмент.
      try {
        var transformations = [
          'w_' + TARGET_W + ',h_' + TARGET_H + ',c_pad,b_black', // размер подложки = целевой
          'du_3.5',
          'ac_none',
          // Оверлей картинки
          'l_' + asset.publicId.replace(/\//g, ':') + ',w_' + TARGET_W + ',h_' + TARGET_H + ',c_pad,b_black,fl_layer_apply,so_0'
        ];

        var convertUrl = 'https://res.cloudinary.com/' + conf.cloudName +
          '/video/upload/' + transformations.join('/') + '/' +
          donorVideoId + '.mp4';

        var vidResp = UrlFetchApp.fetch(convertUrl, { muteHttpExceptions: true });
        if (vidResp.getResponseCode() !== 200) {
          throw new Error('Convert failed ' + vidResp.getResponseCode());
        }

        var blob = vidResp.getBlob().setName('segment_' + asset.index + '.mp4');
        var res = uploadToCloudinary_(blob, 'video');
        Logger.log('-> Segment created: ' + res.publicId);

        return { type: 'video', publicId: res.publicId, width: TARGET_W, height: TARGET_H, index: asset.index };

      } catch (e) {
        Logger.log('Ошибка создания сегмента: ' + e);
        return asset;
      }
    }
  });

  // 3) Склеиваем
  var baseAsset = videoAssets[0];
  var appendAssets = videoAssets.slice(1);

  // База не трансформируется по размеру, она уже эталон (или к ней применяем c_pad если она отличается?)
  // Если первое видео и есть TARGET, то ок. Если TARGET взят из другого video...
  // Для простоты считаем первое видео базой.

  var transformations = [];
  transformations.push('w_' + TARGET_W + ',h_' + TARGET_H + ',c_pad,b_black'); // Гарантируем canvas

  appendAssets.forEach(function (seg) {
    if (seg.type === 'video') {
      var safe = seg.publicId.replace(/\//g, ':');
      // fl_splice ПЕРЕД l_video, и размеры
      var layerTrans = 'w_' + TARGET_W + ',h_' + TARGET_H + ',c_pad,b_black';
      transformations.push('fl_splice,l_video:' + safe);
      transformations.push(layerTrans);
      transformations.push('fl_layer_apply');
    }
  });

  var finalUrl = 'https://res.cloudinary.com/' + conf.cloudName +
    '/video/upload/' + transformations.join('/') + '/' +
    baseAsset.publicId + '.mp4';

  Logger.log('Final Slideshow URL: ' + finalUrl);

  // 4) Проверяем готовность
  var maxWait = 120;
  var deadline = Date.now() + maxWait * 1000;

  while (Date.now() < deadline) {
    var code = cloudinaryPingDelivery_(finalUrl);
    if (code === 200 || code === 206) {
      updateCloudinaryUsage_();
      return finalUrl;
    }
    if (code === 400) throw new Error('Final splice 400 error');
    Utilities.sleep(5000);
  }

  throw new Error('Timeout waiting for slideshow');
}


function buildCloudinaryDeliveryUrl_(cloudName, publicId) {
  var safePublicId = String(publicId).split('/').map(encodeURIComponent).join('/');
  return 'https://res.cloudinary.com/' + cloudName + '/video/upload/' + safePublicId + '.mp4';
}
// === Cloudinary: ожидание готовности (не дольше 60 сек) ===

function cloudinaryAdminGetResource_(resourceType, type, publicId) {
  var conf = getCloudinaryConfig_();
  var url = 'https://api.cloudinary.com/v1_1/' + conf.cloudName +
    '/resources/' + resourceType + '/' + type + '/' + encodeURIComponent(publicId);

  var auth = 'Basic ' + Utilities.base64Encode(conf.apiKey + ':' + conf.apiSecret);

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: auth }
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText();

  Logger.log('Cloudinary Admin GET ' + code + ': ' + body);

  if (code === 404) return null; // ещё не создано — это ок
  if (code < 200 || code >= 300) {
    throw new Error('Cloudinary Admin GET ' + code + ': ' + body);
  }

  return JSON.parse(body);
}

function buildCloudinaryVideoUrl_(publicId, version) {
  var conf = getCloudinaryConfig_();
  var vPart = version ? ('v' + version + '/') : '';
  // mp4 для Postmypost удобнее всего
  return 'https://res.cloudinary.com/' + conf.cloudName + '/video/upload/' + vPart + publicId + '.mp4';
}

function cloudinaryPingDelivery_(url) {
  // чтобы не скачивать весь файл — дергаем 1 байт
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Range: 'bytes=0-0' }
  });
  var code = resp.getResponseCode();
  // 206 Partial Content — идеальный успех для Range
  return code;
}

function waitCloudinaryVideoReady_(publicId, maxSeconds) {
  var conf = getCloudinaryConfig_();
  var deadline = Date.now() + (maxSeconds * 1000);
  var attempt = 0;

  // Строим URL напрямую — для create_slideshow это единственный надёжный способ
  var directUrl = 'https://res.cloudinary.com/' + conf.cloudName + '/video/upload/' + publicId + '.mp4';

  while (Date.now() < deadline) {
    attempt++;
    Logger.log('Cloudinary: жду готовности слайдшоу (' + attempt + ') public_id=' + publicId);

    // Пробуем напрямую пинговать URL (create_slideshow не появляется в Admin API до завершения)
    var pingCode = cloudinaryPingDelivery_(directUrl);
    Logger.log('Cloudinary delivery ping code=' + pingCode + ' url=' + directUrl);

    if (pingCode === 200 || pingCode === 206) {
      Logger.log('Cloudinary slideshow готов!');
      return directUrl;
    }

    // Если 423 (Locked) — значит ещё обрабатывается, это нормально
    if (pingCode === 423) {
      Logger.log('Cloudinary: ресурс ещё обрабатывается (423 Locked)');
    }

    Utilities.sleep(5000); // каждые 5 сек
  }

  throw new Error(
    'Cloudinary: слайдшоу ещё в обработке (>' + maxSeconds + ' сек). public_id=' + publicId
  );
}

// Функция cloudinaryAdminGetResource_ определена выше (строка 594)

function waitCloudinaryResourceReady_(publicId, resourceType, fallbackUrl) {
  var maxAttempts = 72;   // 72 * 5с = 360 секунд = 6 минут
  var delayMs = 5000;

  for (var i = 1; i <= maxAttempts; i++) {
    // 1) Пробуем Admin API (надежнее всего)
    try {
      var res = cloudinaryAdminGetResource_(resourceType, 'upload', publicId);

      if (res) {
        var st = res.status || '';
        Logger.log('Cloudinary Admin status: ' + st + ' (attempt ' + i + '/' + maxAttempts + ')');

        // Как только есть secure_url — возвращаем
        if (res.secure_url || res.url) {
          return res.secure_url || res.url;
        }
      }
    } catch (e) {
      // если тут вылезет 401/403 — увидишь в логе (см. правку ниже)
      Logger.log('Cloudinary Admin check error: ' + e);
    }

    // 2) Легкий пинг delivery URL без скачивания всего видео:
    // Range: bytes=0-0 обычно вернёт 206, если файл уже существует
    try {
      var ping = UrlFetchApp.fetch(fallbackUrl, {
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { Range: 'bytes=0-0' }
      });

      var code = ping.getResponseCode();
      if (code === 206 || code === 200) {
        return fallbackUrl;
      }
      Logger.log('Cloudinary delivery ping code=' + code + ' (attempt ' + i + '/' + maxAttempts + ')');
    } catch (e2) {
      Logger.log('Cloudinary delivery ping error: ' + e2);
    }

    Logger.log('Cloudinary: жду готовности слайдшоу (' + i + '/' + maxAttempts + ') public_id=' + publicId);
    Utilities.sleep(delayMs);
  }

  throw new Error('Cloudinary: не дождался готовности слайдшоу за ~6 минут (public_id=' + publicId + '). URL: ' + fallbackUrl);
}

function sha1Hex_(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    str
  );
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    if (b < 0) b += 256;
    var h = b.toString(16);
    if (h.length === 1) h = '0' + h;
    out += h;
  }
  return out;
}


function parseScheduleDate_(dateVal) {
  if (!dateVal) {
    return { postAt: null, isImmediate: true };
  }

  // Если это уже объект Date (Google Sheets иногда сам парсит)
  if (Object.prototype.toString.call(dateVal) === '[object Date]') {
    var now = new Date();
    // Если дата (включая время) в прошлом или "сейчас" -> немедленно
    if (dateVal <= now) {
      return { postAt: null, isImmediate: true };
    }
    // Иначе это будущее
    var iso = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
    return { postAt: iso, isImmediate: false };
  }

  // Если это строка, пробуем распарсить
  var str = String(dateVal).trim();
  if (!str) return { postAt: null, isImmediate: true };

  // Ожидаемый формат: "DD.MM.YYYY HH:mm" или "DD.MM.YYYY"
  var regex = /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{1,2}))?$/;
  var match = str.match(regex);

  if (!match) {
    return { error: 'Неверный формат даты в ' + CONFIG.SCHEDULE_DATE_CELL + '. Ожидается: ДД.ММ.ГГГГ ЧЧ:мм' };
  }

  var day = parseInt(match[1], 10);
  var month = parseInt(match[2], 10) - 1; 
  var year = parseInt(match[3], 10);
  var hour = match[4] ? parseInt(match[4], 10) : 0;
  var minute = match[5] ? parseInt(match[5], 10) : 0;

  var date = new Date(year, month, day, hour, minute, 0);
  var now = new Date();

  // Валидация корректности даты JS (например 32.01)
  if (date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
     return { error: 'Некорректная дата (такого дня не существует).' };
  }

  if (date <= now) {
    return { postAt: null, isImmediate: true };
  }

  var iso = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
  return { postAt: iso, isImmediate: false };
}


function publishToPostmypostFromSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  var row = CONFIG.OUTPUT_ROW;

  // --- SCHEDULING LOGIC START ---
  var scheduleVal = sheet.getRange(CONFIG.SCHEDULE_DATE_CELL).getValue();
  var scheduleInfo = parseScheduleDate_(scheduleVal);

  if (scheduleInfo.error) {
    setInfo_(scheduleInfo.error);
    return; // Abort if date is invalid
  }
  // --- SCHEDULING LOGIC END ---

  var caption = sheet.getRange(row, CONFIG.COL_CAPTION).getValue();
  // Читаем ссылку из A15 (CONFIG.LINK_CELL)
  var externalLink = sheet.getRange(CONFIG.LINK_CELL).getValue();

  var mediaRaw = sheet.getRange(row, CONFIG.COL_MEDIA_URLS).getValue();
  var mediaUrls = String(mediaRaw || '')
    .split(/\s+|,|\n/)
    .map(function (u) { return u.trim(); })
    .filter(function (u) { return u; });

  if (!mediaUrls.length) {
    throw new Error('В ячейке D' + row + ' нет ссылок на медиа для поста');
  }

  var payload = buildPostmypostPublicationRequest_(mediaUrls, caption, {
    link: externalLink,
    postAt: scheduleInfo.postAt // Pass scheduled time (or null)
  });

  var result = callPostmypost_(payload);

  var pubId = result && (result.id || (result.data && result.data.id));
  if (pubId) {
    Logger.log('Массовая публикация, publication_id=' + pubId);
  }

  if (scheduleInfo.postAt) {
    // Format for display
    setInfo_('Запланировано на ' + scheduleVal);
  } else {
    setInfo_('Опубликовано во все включённые соцсети');
  }
}


function repostSingleSocialFromSheet_(col) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  var layout = CONFIG.SOCIAL_LAYOUT;

  var networkName = sheet.getRange(layout.HEADER_ROW, col).getDisplayValue() || ('столбец ' + col);

  var caption = sheet.getRange(CONFIG.OUTPUT_ROW, CONFIG.COL_CAPTION).getValue();
  // Ссылка для Pinterest (или других, если понадобится)
  var externalLink = sheet.getRange(CONFIG.LINK_CELL).getValue();


  var mediaRaw = sheet.getRange(CONFIG.OUTPUT_ROW, CONFIG.COL_MEDIA_URLS).getValue();
  var mediaUrls = String(mediaRaw || '')
    .split(/\s+|,|\n/)
    .map(function (u) { return u.trim(); })
    .filter(function (u) { return u; });

  if (!mediaUrls.length) {
    setInfo_(
      'Ошибка перевыкладки в ' + networkName +
      ': нет ссылок на медиа в D' + CONFIG.OUTPUT_ROW
    );
    return;
  }

  try {
    setInfo_('Перевыкладка в ' + networkName + ' запущена');

    var payload = buildPostmypostPublicationRequest_(mediaUrls, caption, {
      onlyColumn: col,
      link: externalLink
    });

    var result = callPostmypost_(payload);
    var pubId =
      result && (result.id || (result.data && result.data.id)) || '';

    if (pubId) {
      Logger.log('Перевыкладка в ' + networkName + ', publication_id=' + pubId);
    }

    setInfo_('Перевыложено в ' + networkName);

  } catch (err) {
    Logger.log('Ошибка перевыкладки в ' + networkName + ': ' + err);
    setInfo_(
      'Ошибка перевыкладки в ' + networkName + ': ' +
      (err && err.message ? err.message : err)
    );
  } finally {
    // Сбрасываем выпадашку "Перевыложить?" обратно
    var defaultValue = CONFIG.REPOST_DEFAULT_VALUE || 'Перевыложить?';
    sheet.getRange(layout.REPOST_FLAG_ROW, col).setValue(defaultValue);
  }
}

/***************************
 * СКАЧАТЬ МЕДИА С CDN В BLOB
 ***************************/
function downloadMediaFromUrl_(url) {
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = resp.getResponseCode();

  if (code !== 200) {
    throw new Error('Не удалось скачать медиа ' + url + ': HTTP ' + code +
      ' | ' + resp.getContentText());
  }

  var blob = resp.getBlob();
  var contentType = blob.getContentType() || 'application/octet-stream';

  // ВАЛИДАЦИЯ: если скачали HTML или текст — это ошибка
  if (contentType.indexOf('text/') === 0 || contentType.indexOf('html') !== -1) {
    throw new Error(
      'По ссылке получен не файл, а страница (' + contentType + '). ' +
      'Убедитесь, что ссылка ведёт прямо на картинку/видео, а не на пост.'
    );
  }

  // Простое имя файла из URL
  var cleanUrl = url.split('?')[0];
  var fileName = cleanUrl.split('/').pop() || 'file';

  return {
    blob: blob,
    contentType: contentType,
    fileName: fileName
  };
}

/***************************
 * ШАГ 1: /upload/init (FILE UPLOAD) → uploadId + S3 params
 ***************************/
function initFileUploadInPostmypost_(token, media) {
  var sizeBytes = media.blob.getBytes().length;

  var payload = {
    project_id: PPM_CONFIG.PROJECT_ID,
    name: media.fileName,
    size: sizeBytes
    // при необходимости можно добавить content_type: media.contentType
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify(payload)
  };

  var resp = UrlFetchApp.fetch(PPM_CONFIG.BASE_URL + '/upload/init', options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code !== 200) {
    var msg = body;
    try {
      var json = JSON.parse(body);
      if (json && json.message) msg = json.message;
    } catch (e) { }
    throw new Error('Postmypost /upload/init error: ' + msg);
  }

  var data;
  try {
    data = JSON.parse(body);
  } catch (e) {
    throw new Error('Не удалось распарсить JSON из /upload/init: ' + e + ' | ' + body);
  }

  Logger.log('upload/init (file) response: ' + body);

  var uploadId = data.id;
  if (!uploadId) {
    throw new Error('В ответе /upload/init нет id загрузки. Ответ: ' + body);
  }

  // поля формы для S3: из массива {key, value} → обычный объект
  var fieldsMap = {};
  if (Array.isArray(data.fields)) {
    data.fields.forEach(function (f) {
      fieldsMap[f.key] = f.value;
    });
  }

  var fileUpload = {
    url: data.action,
    fields: fieldsMap
  };

  if (!fileUpload.url) {
    throw new Error('В ответе /upload/init нет action (S3 URL). Ответ: ' + body);
  }

  return {
    uploadId: uploadId,
    fileUpload: fileUpload
  };
}

/***************************
 * ЗАГРУЗКА МЕДИА В POSTMYPOST ЧЕРЕЗ BLOB
 * Шаги на каждый URL:
 * 1) CDN → Blob
 * 2) /upload/init (file) → uploadId + S3 params
 * 3) uploadBlobToS3_ → физически заливаем на S3
 * 4) /upload/complete → говорим Postmypost «я всё залил»
 * 5) /upload/status → ждём готовности и достаём настоящий file_id
 ***************************/
function uploadMediaToPostmypostByUrls_(urls) {
  var token = getPostmypostToken_();
  var fileIds = [];

  urls.forEach(function (url) {
    if (!url) return;

    // 1) качаем файл с CDN в Blob
    var media = downloadMediaFromUrl_(url);

    // 2) /upload/init (file) → uploadId + { url, fields }
    var initResult = initFileUploadInPostmypost_(token, media);
    var uploadId = initResult.uploadId;
    var fileUpload = initResult.fileUpload;

    // 3) шлём Blob в S3
    uploadBlobToS3_(media, fileUpload);

    // 4) фиксируем загрузку в Postmypost
    completeFileUploadInPostmypost_(token, uploadId);

    // 5) ждём, пока Postmypost обработает файл, и ВЫТАСКИВАЕМ file_id
    var fileId = waitFileReadyFromUpload_(token, uploadId);
    if (fileId) {
      fileIds.push(fileId);
    }
  });

  if (!fileIds.length) {
    throw new Error('Не удалось получить ни одного file_id из Postmypost.');
  }

  Logger.log('fileIds для публикации: ' + JSON.stringify(fileIds));
  return fileIds;
}


/***************************
 * ШАГ 2: ОТПРАВИТЬ BLOB НА S3 ПО ДАННЫМ ИЗ file_upload
 ***************************/
function uploadBlobToS3_(media, fileUpload) {
  var fields = fileUpload.fields;
  if (!fields) {
    throw new Error('В объекте file_upload нет поля fields (параметры формы для S3).');
  }

  // Формируем multipart/form-data без ручного boundary:
  // Apps Script сам сделает boundary, если payload — объект + Blob.
  var payload = {};

  Object.keys(fields).forEach(function (key) {
    payload[key] = fields[key];
  });

  // Ключ "file" обычно используется S3 для самого файла
  payload.file = media.blob;

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
    // contentType намеренно НЕ задаём — тогда будет multipart/form-data
  };

  var resp = UrlFetchApp.fetch(fileUpload.url, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  // S3 обычно отвечает 204 или 201; иногда может быть 200
  if (code !== 204 && code !== 201 && code !== 200) {
    throw new Error('Ошибка при загрузке файла в S3: HTTP ' + code + ' | ' + body);
  }
}

/***************************
 * ШАГ 3: /upload/complete → просто фиксируем загрузку
 ***************************/
function completeFileUploadInPostmypost_(token, uploadId) {
  var url = PPM_CONFIG.BASE_URL + '/upload/complete?id=' + encodeURIComponent(uploadId);

  var options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token
    }
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code !== 200) {
    throw new Error('Postmypost /upload/complete ' + code + ': ' + body);
  }

  Logger.log('upload/complete response: ' + body);
}

/***************************
 * ШАГ 4: /upload/status?id=uploadId → ждём и достаём file_id
 ***************************/
function waitFileReadyFromUpload_(token, uploadId) {
  var statusUrl = PPM_CONFIG.BASE_URL + '/upload/status?id=' + encodeURIComponent(uploadId);

  var maxAttempts = 15;   // до ~30 секунд ожидания
  var delayMs = 2000;     // 2 сек между запросами
  var lastBody = '';

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var resp = UrlFetchApp.fetch(statusUrl, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' + token
      }
    });

    var code = resp.getResponseCode();
    var body = resp.getContentText();
    lastBody = body;

    Logger.log('upload/status attempt ' + attempt + ' (code ' + code + '): ' + body);

    if (code !== 200) {
      throw new Error('Postmypost /upload/status ' + code + ': ' + body);
    }

    var data;
    try {
      data = JSON.parse(body);
    } catch (e) {
      throw new Error('Не удалось распарсить JSON из /upload/status: ' + e + ' | ' + body);
    }

    var status = data.status; // 1 / 2 / 3 / ... или строки

    // Файл готов
    if (status === 1 || status === 'COMPLETED') {
      var fileId = null;

      if (data.file_id) {
        fileId = data.file_id;
      } else if (data.file && data.file.id) {
        fileId = data.file.id;
      } else if (data.files && data.files.length && data.files[0].id) {
        fileId = data.files[0].id;
      } else if (data.data && data.data.file && data.data.file.id) {
        fileId = data.data.file.id;
      } else if (data.id) {
        // fallback — если API отдает просто {id: <fileId>, status: 1}
        fileId = data.id;
      }

      if (!fileId) {
        throw new Error(
          'Файл готов (status=' + status +
          '), но file_id не найден. Ответ: ' + body
        );
      }

      Logger.log('Готовый file_id ' + fileId + ' для uploadId ' + uploadId);
      return fileId;
    }

    // Ошибка обработки
    if (status === 2 || status === 'ERROR') {
      throw new Error('Ошибка обработки файла в Postmypost: ' + body);
    }

    // Любой другой статус — ещё обрабатывается
    Utilities.sleep(delayMs);
  }

  throw new Error(
    'Не удалось дождаться готовности файла (uploadId=' +
    uploadId +
    '). Последний ответ /upload/status: ' +
    lastBody
  );
}

/***************************
 * ОТПРАВКА ПУБЛИКАЦИИ В POSTMYPOST (/publications)
 ***************************/
function callPostmypost_(payload) {
  var token = getPostmypostToken_();

  Logger.log('Publications payload: ' + JSON.stringify(payload));

  var options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify(payload)
  };

  var resp = UrlFetchApp.fetch(PPM_CONFIG.BASE_URL + '/publications', options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Postmypost /publications ' + code + ': ' + body);
  }

  var data;
  try {
    data = JSON.parse(body);
  } catch (e) {
    throw new Error('Не удалось распарсить ответ Postmypost: ' + e + ' | ' + body);
  }

  return data;
}

/***************************
 * ТОКЕН POSTMYPOST ИЗ SCRIPT PROPERTIES
 ***************************/
function getPostmypostToken_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(PPM_CONFIG.TOKEN_PROPERTY);
  if (!token) {
    throw new Error('В Script Properties не найден токен Postmypost с ключом ' + PPM_CONFIG.TOKEN_PROPERTY);
  }
  return token;
}


/***************************
 * ОБРАБОТЧИК УСТАНАВЛИВАЕМОГО onEdit
 ***************************/
function onEditHandler(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  var a1 = e.range.getA1Notation();
  var value = e.value;
  var layout = CONFIG.SOCIAL_LAYOUT;
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // 1) Получить последний пост
  if (a1 === CONFIG.DROPDOWN_CELL_FETCH &&
    value === CONFIG.FETCH_TRIGGER_VALUE) {
    handleFetchInstagram_();
    return;
  }

  // 1b) Получить пост по ссылке (E1)
  if (a1 === CONFIG.DROPDOWN_CELL_FETCH_BY_LINK &&
    value === CONFIG.FETCH_BY_LINK_TRIGGER_VALUE) {
    handleFetchInstagramByLink_();
    return;
  }

  // 2) Адаптировать текст под все соцсети (общая кнопка в D1)
  if (a1 === CONFIG.DROPDOWN_CELL_ADAPT &&
    value === CONFIG.ADAPT_TRIGGER_VALUE) {
    handleAdaptText_();
    return;
  }

  // 3) Выложить пост во все включённые соцсети (A1)
  if (a1 === CONFIG.DROPDOWN_CELL_PUBLISH &&
    value === CONFIG.PUBLISH_TRIGGER_VALUE) {
    handlePublishPost_();
    return;
  }

  // 4) ЛОКАЛЬНОЕ переписывание текста для одной соцсети
  // строка "Переписать текст?" (REWRITE_FLAG_ROW)
  if (row === layout.REWRITE_FLAG_ROW &&
    String(value || '').toLowerCase() === 'да') {

    handleSingleRewriteFlag_(sheet, col);
    return;
  }

  // 5) ЛОКАЛЬНАЯ перевыкладка для одной соцсети
  // строка "Перевыложить?" (REPOST_FLAG_ROW)
  if (row === layout.REPOST_FLAG_ROW &&
    String(value || '').toLowerCase() === 'да') {

    // Вся логика + логирование/сброс выпадашки уже внутри функции
    repostSingleSocialFromSheet_(col);
    return;
  }
}

function handleFetchInstagramByLink_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  var link = String(
    sheet.getRange(CONFIG.INSTAGRAM_LINK_CELL).getValue() || ''
  ).trim();

  if (!link) {
    throw new Error(
      'В ячейке ' + CONFIG.INSTAGRAM_LINK_CELL + ' нет ссылки на пост Instagram'
    );
  }

  try {
    setInfo_('Забираю пост из Instagram по ссылке');

    var shortcode = extractInstagramShortcodeFromUrl_(link);
    var info = fetchInstagramPostDataByShortcode_(shortcode);

    fillOutputRowFromInstagramInfo_(sheet, info, false);

    setInfo_('Пост из Instagram по ссылке обновлён');
  } catch (err) {
    Logger.log('Ошибка в handleFetchInstagramByLink_: ' + err);
    setInfo_(
      'Ошибка при получении поста по ссылке: ' +
      (err && err.message ? err.message : err)
    );
  } finally {
    sheet.getRange(CONFIG.DROPDOWN_CELL_FETCH_BY_LINK)
      .setValue(CONFIG.FETCH_BY_LINK_DEFAULT_VALUE);
  }
}

function handleSingleRewriteFlag_(sheet, col) {
  var layout = CONFIG.SOCIAL_LAYOUT;
  var name = sheet.getRange(layout.HEADER_ROW, col).getDisplayValue() || ('столбец ' + col);

  try {
    setInfo_('Адаптация текста для ' + name + ' запущена');
    adaptTextForSingleNetwork_(col);
    setInfo_('Адаптация текста для ' + name + ' завершена');
  } catch (err) {
    Logger.log('Ошибка одиночной адаптации для столбца ' + col + ': ' + err);
    sheet.getRange(layout.OUTPUT_ROW, col)
      .setValue('Ошибка адаптации: ' + (err && err.message ? err.message : err));
    setInfo_(
      'Ошибка адаптации текста для ' + name + ': ' +
      (err && err.message ? err.message : err)
    );
  } finally {
    sheet.getRange(layout.REWRITE_FLAG_ROW, col)
      .setValue(CONFIG.REWRITE_DEFAULT_VALUE || 'Переписать текст?');
  }
}

/***************************
 * ОБРАБОТЧИК: ВЫЛОЖИТЬ ПОСТ В СОЦСЕТИ
 ***************************/
function handlePublishPost_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  try {
    setInfo_('Публикация во все включённые соцсети запущена');
    publishToPostmypostFromSheet_();
  } catch (err) {
    Logger.log('Ошибка в handlePublishPost_: ' + err);
    setInfo_('Ошибка публикации: ' + (err && err.message ? err.message : err));
  } finally {
    sheet.getRange(CONFIG.DROPDOWN_CELL_PUBLISH)
      .setValue(CONFIG.PUBLISH_DEFAULT_VALUE);
  }
}

/***************************
 * АВТОМАТИЧЕСКАЯ ВЫКЛАДКА КАЖДЫЙ ЧАС
 * 1) Проверяем, включена ли автоворонка в C1
 * 2) Тянем последний пост из Инсты
 * 3) Сравниваем его postKey с тем, что лежит в F3
 * 4) Если новый → обновляем строку 3 и вызываем publishToPostmypostFromSheet_()
 ***************************/
function autoCheckAndCrosspost_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('Автовыкладка: лист ' + CONFIG.SHEET_NAME + ' не найден');
    return;
  }

  // Проверяем переключатель в C1
  var autoVal = sheet.getRange(CONFIG.DROPDOWN_CELL_AUTO).getDisplayValue();
  if (String(autoVal || '').trim() !== CONFIG.AUTO_ON_VALUE) {
    Logger.log('Автовыкладка: выключена (C1="' + autoVal + '"), выходим.');
    return;
  }

  var row = CONFIG.OUTPUT_ROW;

  try {
    // 1) Берём ранее сохранённый ID поста из F3
    var lastKey = String(
      sheet.getRange(row, CONFIG.COL_POST_ID).getValue() || ''
    ).trim();

    // 2) Тянем самый свежий пост из Инсты
    var latest = fetchLatestInstagramPostData_();
    var newKey = String(latest.postKey || '').trim();

    if (!newKey) {
      setInfo_('Автовыкладка: не удалось получить ID поста из Instagram');
      Logger.log('Автовыкладка: пустой postKey у последнего поста');
      return;
    }

    // 3) Сравниваем — если такой же, выходим
    if (newKey === lastKey) {
      Logger.log('Автовыкладка: новых постов нет (postKey=' + newKey + ')');
      return;
    }

    // Есть новый пост
    setInfo_('Автовыкладка: найден новый пост, обновляю данные из Instagram');
    fillOutputRowFromInstagramInfo_(sheet, latest, true);

    // 5) Адаптация текста для всех включённых соцсетей
    setInfo_('Автовыкладка: адаптирую текст для включённых соцсетей');
    adaptLastInstagramPostForNetworks_();

    // 6) Публикация во все включённые соцсети
    setInfo_('Автовыкладка: публикую пост во все включённые соцсети');
    publishToPostmypostFromSheet_();

    // Финальный статус
    setInfo_('Автовыкладка: пост опубликован');

  } catch (err) {
    Logger.log('Автовыкладка: ошибка ' + err);
    setInfo_(
      'Автовыкладка: ошибка — ' +
      (err && err.message ? err.message : err)
    );
  }
}

/***************************
 * ОБРАБОТЧИК: ЗАБРАТЬ ПОСЛЕДНИЙ ПОСТ
 ***************************/
function handleFetchInstagram_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  try {
    setInfo_('Забираю последний пост из Instagram');
    updateLastInstagramPost_();
    setInfo_('Последний пост из Instagram обновлён');
  } catch (err) {
    Logger.log('Ошибка в handleFetchInstagram_: ' + err);
    setInfo_('Ошибка при получении поста: ' + (err && err.message ? err.message : err));
  } finally {
    sheet.getRange(CONFIG.DROPDOWN_CELL_FETCH)
      .setValue(CONFIG.FETCH_DEFAULT_VALUE);
  }
}

function extractInstagramShortcodeFromUrl_(url) {
  if (!url) {
    throw new Error('Ссылка на пост Instagram пуста');
  }

  var trimmed = String(url).trim();

  // Если дали просто сам shortcode
  if (/^[A-Za-z0-9_\-]+$/.test(trimmed)) {
    return trimmed;
  }

  // Ищем `/p/SHORT/`, `/reel/SHORT/`, `/tv/SHORT/`
  var match = trimmed.match(/instagram\.com\/(?:reel|p|tv)\/([A-Za-z0-9_\-]+)/i);
  if (match && match[1]) {
    return match[1];
  }

  throw new Error('Не удалось выделить shortcode из ссылки: ' + url);
}

function fetchInstagramPostDataByShortcode_(shortcode) {
  var sc = String(shortcode || '').trim();
  if (!sc) {
    throw new Error('Shortcode пуст — нечего запрашивать');
  }

  var payload = {
    shortcode: sc
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      'x-rapidapi-host': CONFIG.RAPIDAPI_HOST,
      'x-rapidapi-key': getRapidApiKey_()
    },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(CONFIG.API_MEDIA_BY_SHORTCODE_ENDPOINT, options);
  updateRapidApiQuota_(resp.getHeaders());
  var code = resp.getResponseCode();

  if (code !== 200) {
    throw new Error(
      'RapidAPI mediaByShortcode вернул код ' +
      code + ': ' + resp.getContentText()
    );
  }

  var data;
  try {
    data = JSON.parse(resp.getContentText());
  } catch (err) {
    throw new Error('Не удалось распарсить JSON mediaByShortcode: ' + err);
  }

  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Пустой или некорректный ответ от mediaByShortcode');
  }

  // В этом эндпоинте весь массив — медиа одного поста
  var samePostItems = data;
  var latestItem = samePostItems[0]; // для превью сгодится первый

  var meta = latestItem.meta || {};
  var username = meta.username || '';
  var realShortcode = meta.shortcode || sc;

  // Все URL'ы медиа
  var mediaUrls = extractMediaUrls_(samePostItems);

  // Определяем тип поста
  var type = determinePostType_(samePostItems);

  // Картинка для превью
  var imageUrl =
    latestItem.pictureUrl ||
    (latestItem.urls &&
      latestItem.urls[0] &&
      latestItem.urls[0].url) ||
    '';

  var caption = meta.title || '';
  var postUrl =
    meta.sourceUrl ||
    (realShortcode ? 'https://www.instagram.com/p/' + realShortcode + '/' : '');

  var postKey = buildPostKey_(meta, realShortcode);

  return {
    imageUrl: imageUrl,
    caption: caption,
    type: type,
    postUrl: postUrl,
    username: username,
    mediaUrls: mediaUrls,
    postKey: postKey
  };
}

function fillOutputRowFromInstagramInfo_(sheet, info, savePostId) {
  var row = CONFIG.OUTPUT_ROW;

  // чистим первые 4 столбца (A:D)
  sheet.getRange(row, 1, 1, 4).clearContent();

  // A: превью
  if (info.imageUrl) {
    var escapedUrl = String(info.imageUrl).replace(/"/g, '""');
    sheet.getRange(row, CONFIG.COL_IMAGE)
      .setFormula('=IMAGE("' + escapedUrl + '")');
  } else if (info.mediaUrls && info.mediaUrls[0]) {
    sheet.getRange(row, CONFIG.COL_IMAGE).setValue(info.mediaUrls[0]);
  }

  // B: текст
  sheet.getRange(row, CONFIG.COL_CAPTION).setValue(info.caption || '');

  // C: тип поста
  sheet.getRange(row, CONFIG.COL_TYPE).setValue(info.type || '');

  // D: медиа-URL’ы (по одному в строке)
  if (info.mediaUrls && info.mediaUrls.length) {
    sheet.getRange(row, CONFIG.COL_MEDIA_URLS)
      .setValue(info.mediaUrls.join('\n'));
  }

  // F: уникальный ID поста (shortcode/ключ) — пишем ТОЛЬКО при автопостинге (или явном флаге)
  if (savePostId) {
    sheet.getRange(row, CONFIG.COL_POST_ID)
      .setValue(info.postKey || '');
  }
}

function updateRapidApiQuota_(headers) {
  if (!headers) return;

  // RapidAPI headers come in lowercase from UrlFetchApp usually, but let's be safe
  // Example headers:
  // x-ratelimit-requests-limit: 1000
  // x-ratelimit-requests-remaining: 978
  // x-ratelimit-requests-reset: 962202 (seconds remaining)

  var limit = headers['x-ratelimit-requests-limit'];
  var remaining = headers['x-ratelimit-requests-remaining'];
  var resetSec = headers['x-ratelimit-requests-reset'];

  // If simple access fails, try case-insensitive search
  if (limit == null || remaining == null || resetSec == null) {
    var lowerKeys = {};
    for (var k in headers) {
      lowerKeys[k.toLowerCase()] = headers[k];
    }
    limit = lowerKeys['x-ratelimit-requests-limit'];
    remaining = lowerKeys['x-ratelimit-requests-remaining'];
    resetSec = lowerKeys['x-ratelimit-requests-reset'];
  }

  if (limit == null || remaining == null) {
    // Not found or different API plan
    return;
  }

  // Calculate human readable reset date
  var resetText = '';
  var sec = parseInt(resetSec, 10);
  if (!isNaN(sec)) {
    var resetDate = new Date(Date.now() + sec * 1000);
    var day = ('0' + resetDate.getDate()).slice(-2);
    var month = ('0' + (resetDate.getMonth() + 1)).slice(-2);
    resetText = 'Обновится ' + day + '.' + month;
  }

  var msg = remaining + ' из ' + limit + '\n' + resetText;

  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) {
    sheet.getRange(CONFIG.RAPIDAPI_QUOTA_CELL).setValue(msg.trim());
  }
}

/***************************
 * ОБРАБОТЧИК: АДАПТИРОВАТЬ ТЕКСТ ПОД СОЦСЕТИ
 ***************************/
function handleAdaptText_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  var layout = CONFIG.SOCIAL_LAYOUT;

  try {
    setInfo_('Адаптация текста для всех соцсетей запущена');
    adaptLastInstagramPostForNetworks_();
    setInfo_('Адаптация текста для всех соцсетей завершена');
  } catch (err) {
    Logger.log('Ошибка в handleAdaptText_: ' + err);

    // старая ошибка под первой соцсетью — можно оставить
    sheet.getRange(layout.OUTPUT_ROW, layout.FIRST_COL)
      .setValue('Ошибка адаптации: ' + (err && err.message ? err.message : err));

    // и дублируем её в Info
    setInfo_('Ошибка адаптации: ' + (err && err.message ? err.message : err));
  } finally {
    sheet.getRange(CONFIG.DROPDOWN_CELL_ADAPT)
      .setValue(CONFIG.ADAPT_DEFAULT_VALUE);
  }
}


/***************************
 * ПОЛУЧЕНИЕ ПОСЛЕДНЕГО ПОСТА ИЗ ИНСТЫ
 ***************************/
function updateLastInstagramPost_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  var info = fetchLatestInstagramPostData_();
  fillOutputRowFromInstagramInfo_(sheet, info, true);
}


/***************************
 * РАБОТА С SCRIPT PROPERTIES
 ***************************/
function getRapidApiKey_() {
  var key = PropertiesService.getScriptProperties()
    .getProperty(PROP_KEYS.RAPIDAPI_KEY);
  if (!key) {
    throw new Error(
      'RapidAPI key не найден в Script Properties. ' +
      'Задай его в Script Properties с именем ' + PROP_KEYS.RAPIDAPI_KEY + '.'
    );
  }
  return key;
}

function getOpenAiApiKey_() {
  var key = PropertiesService.getScriptProperties()
    .getProperty(PROP_KEYS.OPENAI_API_KEY);
  if (!key) {
    throw new Error(
      'OpenAI API key не найден в Script Properties. ' +
      'Задай его в Script Properties с именем ' + PROP_KEYS.OPENAI_API_KEY + '.'
    );
  }
  return key;
}

/***************************
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ОБРАБОТКИ ДАННЫХ
 ***************************/

/**
 * Извлекает URL медиа-файлов из массива элементов Instagram поста
 */
function extractMediaUrls_(items) {
  return items
    .map(function (item) {
      if (item.urls && item.urls[0] && item.urls[0].url) {
        return item.urls[0].url;
      }
      return null;
    })
    .filter(function (u) { return !!u; });
}

/**
 * Определяет тип поста на основе расширений файлов и количества элементов
 */
function determinePostType_(items) {
  var videoCount = 0;
  var imageCount = 0;

  items.forEach(function (item) {
    if (item.urls && item.urls[0] && item.urls[0].extension) {
      var ext = String(item.urls[0].extension).toLowerCase();
      if (ext === 'mp4' || ext === 'mov' || ext === 'avi') {
        videoCount++;
      } else {
        imageCount++;
      }
    } else {
      // Если расширения нет, но item пришел, считаем по умолчанию картинкой (безопасный фоллбэк)
      imageCount++;
    }
  });

  var total = videoCount + imageCount;

  if (total <= 1) {
    if (videoCount > 0) return 'Reels';
    return 'Фото';
  } else {
    // Карусели
    if (videoCount > 0 && imageCount > 0) {
      return 'Карусель фото + видео';
    } else if (videoCount > 0) {
      return 'Карусель видео';
    } else {
      return 'Карусель фото';
    }
  }
}

/**
 * Создаёт уникальный ключ поста для сравнения/идентификации
 */
function buildPostKey_(meta, shortcode) {
  return (meta.id && String(meta.id)) ||
    (shortcode ? String(shortcode) : '') ||
    (meta.takenAt ? 'takenAt_' + meta.takenAt : '');
}

/**
 * Парсит ответ OpenAI и извлекает title и text
 */
function parseTitleAndText_(raw) {
  var title = '';
  var text = '';

  // пробуем разобрать как JSON {"title":"...","text":"..."}
  try {
    var parsed = JSON.parse(raw);
    title = parsed.title || '';
    text = parsed.text || parsed.body || '';
  } catch (jsonErr) {
    // запасной план: первая непустая строка — заголовок, остальное — текст
    var lines = String(raw).split('\n');
    for (var j = 0; j < lines.length; j++) {
      if (lines[j].trim()) {
        title = lines[j].replace(/^Заголовок\s*[:\-–]\s*/i, '').trim();
        text = lines.slice(j + 1).join('\n').trim();
        break;
      }
    }
    if (!text) {
      text = String(raw).trim();
    }
  }

  return { title: title, text: text };
}

/***************************
 * RAPIDAPI: ИНСТА → ДАННЫЕ ПО ПОСТУ
 ***************************/
function fetchLatestInstagramPostData_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');
  // Базовый URL профиля из конфига
  var sourceUrl = String(
    sheet.getRange(CONFIG.ROW_INSTAGRAM_URL, CONFIG.COL_INSTAGRAM_URL).getValue() || ''
  ).trim();
  // Делаем "анти-кеш" версию с другим регистром ника
  var urlForApi = jitterInstagramUsernameCase_(sourceUrl);

  var payload = {
    url: urlForApi
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      'x-rapidapi-host': CONFIG.RAPIDAPI_HOST,
      'x-rapidapi-key': getRapidApiKey_()
    },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(CONFIG.API_ENDPOINT, options);
  updateRapidApiQuota_(resp.getHeaders());
  var code = resp.getResponseCode();

  if (code !== 200) {
    throw new Error('RapidAPI вернул код ' + code + ': ' + resp.getContentText());
  }

  var data;
  try {
    data = JSON.parse(resp.getContentText());
  } catch (err) {
    throw new Error('Не удалось распарсить JSON: ' + err);
  }

  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Пустой или некорректный ответ от API');
  }

  // Ищем самый свежий пост по meta.takenAt
  var latestItem = data.reduce(function (best, item) {
    if (!item || !item.meta || typeof item.meta.takenAt !== 'number') return best;
    if (!best) return item;
    return item.meta.takenAt > best.meta.takenAt ? item : best;
  }, null) || data[0];

  var meta = latestItem.meta || {};
  var shortcode = meta.shortcode;
  var username = meta.username || '';

  // Собираем все элементы с этим же shortcode (для карусели)
  var samePostItems = shortcode
    ? data.filter(function (item) {
      return item.meta && item.meta.shortcode === shortcode;
    })
    : [latestItem];

  // Все основные media-URL'ы
  var mediaUrls = extractMediaUrls_(samePostItems);

  // Тип поста
  var type = determinePostType_(samePostItems);

  // Картинка для превью
  var imageUrl =
    latestItem.pictureUrl ||
    (latestItem.urls &&
      latestItem.urls[0] &&
      latestItem.urls[0].url) ||
    '';

  var caption = meta.title || '';
  var postUrl = shortcode ? 'https://www.instagram.com/p/' + shortcode + '/' : '';

  // 🔑 Уникальный ключ поста для сравнения
  var postKey = buildPostKey_(meta, shortcode);

  return {
    imageUrl: imageUrl,
    caption: caption,
    type: type,
    postUrl: postUrl,
    username: username,
    mediaUrls: mediaUrls,
    postKey: postKey
  };
}


/***************************
 * АДАПТАЦИЯ ТЕКСТА ПОД СОЦСЕТИ
 ***************************/
function adaptLastInstagramPostForNetworks_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  // базовый текст — исходная подпись из Инсты (B3)
  var baseText = sheet.getRange(CONFIG.OUTPUT_ROW, CONFIG.COL_CAPTION).getValue();
  if (!baseText) {
    throw new Error('Исходный текст поста (ячейка B' + CONFIG.OUTPUT_ROW + ') пустой');
  }

  var layout = CONFIG.SOCIAL_LAYOUT;

  // читаем все соцсети и берём только включённые + с моделью и промптом
  var socials = getSocialRowsFromSheet_().filter(function (s) {
    return s.enabled && s.model && s.prompt;
  });

  if (!socials.length) {
    Logger.log('Нет ни одной включённой соцсети для адаптации');
    return;
  }

  var apiKey = getOpenAiApiKey_();
  var url = 'https://api.openai.com/v1/chat/completions';

  // собираем запросы к OpenAI
  var jobs = socials.map(function (s) {
    var messages = [
      {
        role: 'system',
        content: getSystemPrompt_(sheet)
      },
      {
        role: 'user',
        content:
          'Платформа: ' + s.name + '.\n' +
          'Задача: ' + s.prompt + '\n\n' +
          'Исходный текст поста:\n"""' + baseText + '"""'
      }
    ];

    return {
      social: s,
      request: {
        url: url,
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: 'Bearer ' + apiKey
        },
        payload: JSON.stringify({
          model: s.model,
          messages: messages
          // temperature не задаём — дефолт (1), как просит модель
        }),
        muteHttpExceptions: true
      }
    };
  });

  // шлём все запросы параллельно
  var requests = jobs.map(function (job) { return job.request; });
  var responses = UrlFetchApp.fetchAll(requests);

  responses.forEach(function (resp, i) {
    var job = jobs[i];
    var s = job.social;
    var col = s.col;

    var titleCell = sheet.getRange(layout.TITLE_ROW, col);
    var textCell = sheet.getRange(layout.OUTPUT_ROW, col);

    try {
      var code = resp.getResponseCode();
      if (code !== 200) {
        titleCell.setValue('');
        textCell.setValue('Ошибка OpenAI ' + code + ': ' + resp.getContentText());
        return;
      }

      var data = JSON.parse(resp.getContentText());
      if (!data.choices || !data.choices[0] || !data.choices[0].message) {
        titleCell.setValue('');
        textCell.setValue('Неожиданный ответ OpenAI');
        return;
      }

      var raw = data.choices[0].message.content;
      var result = parseTitleAndText_(raw);

      titleCell.setValue(result.title);
      textCell.setValue(result.text);

    } catch (err) {
      titleCell.setValue('');
      textCell.setValue('Ошибка: ' + (err && err.message ? err.message : err));
      Logger.log('Ошибка адаптации для ' + s.name + ': ' + err);
    }
  });
}

function adaptTextForSingleNetwork_(col) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Лист ' + CONFIG.SHEET_NAME + ' не найден');

  // базовый текст — подпись из Инсты (B3)
  var baseText = sheet.getRange(CONFIG.OUTPUT_ROW, CONFIG.COL_CAPTION).getValue();
  if (!baseText) {
    throw new Error('Исходный текст поста (ячейка B' + CONFIG.OUTPUT_ROW + ') пустой');
  }

  var layout = CONFIG.SOCIAL_LAYOUT;

  // читаем настройки для конкретного столбца
  var name = String(sheet.getRange(layout.HEADER_ROW, col).getValue() || '').trim();
  if (!name) {
    throw new Error('В строке ' + layout.HEADER_ROW +
      ' нет названия соцсети для столбца ' + col);
  }

  var model = String(sheet.getRange(layout.MODEL_ROW, col).getValue() || '').trim();
  if (!model) {
    throw new Error('Для соцсети "' + name +
      '" не указана модель (строка ' + layout.MODEL_ROW + ')');
  }

  var prompt = String(sheet.getRange(layout.PROMPT_ROW, col).getValue() || '').trim();
  if (!prompt) {
    throw new Error('Для соцсети "' + name +
      '" не задан промпт (строка ' + layout.PROMPT_ROW + ')');
  }

  var apiKey = getOpenAiApiKey_();
  var url = 'https://api.openai.com/v1/chat/completions';

  var messages = [
    {
      role: 'system',
      content: getSystemPrompt_(sheet)
    },
    {
      role: 'user',
      content:
        'Платформа: ' + name + '.\n' +
        'Задача: ' + prompt + '\n\n' +
        'Исходный текст поста:\n"""' + baseText + '"""'
    }
  ];

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify({
      model: model,
      messages: messages   // без temperature — использует дефолт 1
    }),
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();

  var titleCell = sheet.getRange(layout.TITLE_ROW, col);
  var textCell = sheet.getRange(layout.OUTPUT_ROW, col);

  if (code !== 200) {
    titleCell.setValue('');
    textCell.setValue('Ошибка OpenAI ' + code + ': ' + resp.getContentText());
    return;
  }

  var data;
  try {
    data = JSON.parse(resp.getContentText());
  } catch (e) {
    titleCell.setValue('');
    textCell.setValue('Не удалось распарсить ответ OpenAI: ' + e);
    return;
  }

  if (!data.choices || !data.choices[0] || !data.choices[0].message) {
    titleCell.setValue('');
    textCell.setValue('Неожиданный ответ OpenAI');
    return;
  }

  var raw = data.choices[0].message.content;
  var result = parseTitleAndText_(raw);

  titleCell.setValue(result.title);
  textCell.setValue(result.text);
}

// === Cloudinary Cleanup Functions ===

/**
 * Public function to clean up Cloudinary assets older than 1 hour.
 * Recommended to run via a Time-driven trigger (e.g., daily).
 */
function cleanCloudinary() {
  Logger.log('Starting Cloudinary cleanup (older than 1 hour)...');

  var now = Date.now();
  var oneHourAgo = new Date(now - 60 * 60 * 1000);

  var resultsImg = runCleanupForType_('image', oneHourAgo);
  var resultsVid = runCleanupForType_('video', oneHourAgo);

  Logger.log('Cloudinary Cleanup Finished. Deleted Images: ' + resultsImg + ', Videos: ' + resultsVid);
  setInfo_('Cloudinary cleaned: ' + resultsImg + ' images, ' + resultsVid + ' videos');

  updateCloudinaryUsage_();
}

/**
 * Helper to run cleanup for a specific resource type.
 */
function runCleanupForType_(resourceType, dateObj) {
  var timestamp = dateObj.toISOString();
  // expression: "resource_type:image AND created_at < '2025-12-25T...'"
  var expression = 'resource_type:' + resourceType + ' AND created_at < "' + timestamp + '"';

  var publicIds = cloudinarySearch_(expression);

  if (!publicIds.length) {
    Logger.log('No old ' + resourceType + ' assets found.');
    return 0;
  }

  Logger.log('Found ' + publicIds.length + ' old ' + resourceType + ' assets. Deleting...');

  var deletedCount = 0;
  var BATCH_SIZE = 50; // Keep batch size moderate

  for (var i = 0; i < publicIds.length; i += BATCH_SIZE) {
    var batch = publicIds.slice(i, i + BATCH_SIZE);
    try {
      cloudinaryDeleteBatch_(batch, resourceType);
      deletedCount += batch.length;
    } catch (e) {
      Logger.log('Error deleting batch ' + resourceType + ': ' + e);
    }
  }
  return deletedCount;
}

/**
 * Search Cloudinary resources using Admin API Search method.
 */
function cloudinarySearch_(expression) {
  var conf = getCloudinaryConfig_();
  var url = 'https://api.cloudinary.com/v1_1/' + conf.cloudName + '/resources/search';

  var auth = 'Basic ' + Utilities.base64Encode(conf.apiKey + ':' + conf.apiSecret);

  var allIds = [];
  var nextCursor = null;

  do {
    var payload = {
      expression: expression,
      max_results: 500,
      fields: ['public_id']
    };
    if (nextCursor) {
      payload.next_cursor = nextCursor;
    }

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: auth },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var txt = resp.getContentText();

    if (code !== 200) {
      throw new Error('Cloudinary Search Error (' + code + '): ' + txt);
    }

    var data = JSON.parse(txt);
    if (data.resources) {
      data.resources.forEach(function (r) {
        allIds.push(r.public_id);
      });
    }

    nextCursor = data.next_cursor || null;

  } while (nextCursor);

  return allIds;
}

/**
 * Delete a batch of resources using Admin API.
 */
function cloudinaryDeleteBatch_(publicIds, resourceType) {
  if (!publicIds || !publicIds.length) return;

  var conf = getCloudinaryConfig_();
  // Admin API Delete Resources: DELETE /resources/:resource_type/upload
  var url = 'https://api.cloudinary.com/v1_1/' + conf.cloudName + '/resources/' + resourceType + '/upload';

  // Construct query string for public_ids array: public_ids[]=id1&public_ids[]=id2...
  var qs = publicIds.map(function (id) {
    return 'public_ids[]=' + encodeURIComponent(id);
  }).join('&');

  var fullUrl = url + '?' + qs;

  var auth = 'Basic ' + Utilities.base64Encode(conf.apiKey + ':' + conf.apiSecret);

  var options = {
    method: 'delete',
    headers: { Authorization: auth },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(fullUrl, options);
  var code = resp.getResponseCode();

  if (code !== 200) {
    throw new Error('Cloudinary Delete Error (' + code + '): ' + resp.getContentText());
  }
  Logger.log('Deleted batch of ' + publicIds.length + ' assets.');
}

function updateCloudinaryUsage_() {
  var conf = getCloudinaryConfig_();
  if (!conf.cloudName || !conf.apiKey || !conf.apiSecret) return;

  var url = 'https://api.cloudinary.com/v1_1/' + conf.cloudName + '/usage';
  var auth = 'Basic ' + Utilities.base64Encode(conf.apiKey + ':' + conf.apiSecret);

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: auth },
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('Cloudinary Usage fetch failed (' + code + '): ' + resp.getContentText());
      return;
    }

    var data = JSON.parse(resp.getContentText());

    // Перевод и пересчет даты
    var msgLines = [];

    // 1. План
    if (data.plan) {
      msgLines.push('План: ' + data.plan);
    }

    // 2. Кредиты
    if (data.credits) {
      var used = data.credits.usage || 0;
      var limit = data.credits.limit || 0;
      var percent = limit > 0 ? Math.round((used / limit) * 100) : 0;
      msgLines.push('Кредитов израсходовано: ' + percent + '% (' + used.toFixed(2) + ' из ' + limit + ')');
    }

    // 3. Дата обновления квоты (30-дневный цикл от CONFIG.CLOUDINARY_REG_DATE)
    if (CONFIG.CLOUDINARY_REG_DATE) {
      var regDate = new Date(CONFIG.CLOUDINARY_REG_DATE);
      if (!isNaN(regDate.getTime())) {
        var now = new Date();
        var diffTime = now.getTime() - regDate.getTime();
        // 30 дней в миллисекундах
        var periodMs = 30 * 24 * 60 * 60 * 1000;

        // Сколько периодов прошло полностью
        var periodsPassed = Math.floor(diffTime / periodMs);
        // Конец текущего периода
        var distinctNextReset = regDate.getTime() + (periodsPassed + 1) * periodMs;

        var daysLeft = Math.ceil((distinctNextReset - now.getTime()) / (1000 * 60 * 60 * 24));
        var resetDateObj = new Date(distinctNextReset);

        var dayStr = ('0' + resetDateObj.getDate()).slice(-2);
        var monthStr = ('0' + (resetDateObj.getMonth() + 1)).slice(-2);

        msgLines.push('Обновится ' + dayStr + '.' + monthStr);
      }
    }

    // Storage убрали по просьбе

    var finalMsg = msgLines.join('\n');

    var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) {
      sheet.getRange(CONFIG.CLOUDINARY_USAGE_CELL).setValue(finalMsg);
    }

  } catch (e) {
    Logger.log('Error updating Cloudinary usage: ' + e);
  }
}