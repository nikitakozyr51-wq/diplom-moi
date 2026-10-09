# Материалы диплома

Основной текущий текст: [working-draft.docx](drafts/working-draft.docx?raw=true) и [working-draft.md](drafts/working-draft.md). Word и Markdown содержат главу 1, начальную российскую часть главы 2 и 16 использованных источников. Для скачивания через GitHub используйте ссылки из [главной страницы](../README.md).

- [Подробный план](plan.md) и [план в Word](plan.docx?raw=true).
- [Порядок работы](work-plan.md), [контекст](project-context.md), [применение скилла](academic-workflow.md).
- [Проверка по ARS](review-2026-10-09.md), [статус готовности](verification-status.md), [журнал ошибок](error-log.md).
- [Протокол аналитической части](analysis/scenario-protocol.md) — предложенный метод, ещё не выполненное исследование.
- [Разбор семи учебных разделов](literature/course/reading-notes.md); [карта прочитанных источников](literature/reading-notes.md).
- [Российские примеры](projects/russian-cases.md) — три группы эпизодов одного разработчика, с границами доказательств.
- [Требования вуза](requirements/requirements.md).
- [sources.csv](sources.csv): 17 источников; [evidence.csv](evidence.csv): 44 утверждения; [карта номеров ссылок](literature/citation-map.csv).
- [Библиография](bibliography.md), [контрольные суммы](literature/file-manifest.json), [проверка структуры PDF](literature/pdf-preflight/).
- [Рабочая глава 1](drafts/chapter-1.docx?raw=true) и [начало главы 2](drafts/chapter-2-russia-start.md). Номера ссылок отдельных глав относятся к общему списку.

`chapter-1-start.*` и `chapter-1-continuation.md` сохранены как ранние фрагменты. Для проверки актуального текста используйте общий `working-draft.*`, а не смешивайте версии.

Оригиналы методических и учебных документов и полные извлечения остаются в облачной среде, вне публичного Git. Их происхождение и контрольные суммы сохранены. Локальный ARS установлен в `.agents/skills/academic-research-suite/`; зависимости — в `/workspace/shared/ars-environment/`. Сам скилл, окружение Python и профили браузера не публикуются как материалы диплома.

Экспорт: сначала `tools/export-plan.py`, затем `tools/export-draft.py`; Python окружения — `/workspace/shared/ars-environment/venv/bin/python`, нужен Pandoc. Проверка — `tools/check-draft.py`; ключ `--require-local-sources` требует оригиналы из облачной среды. Файлы `build/` содержат воспроизводимые отчёты и справочные стили, а не окончательное заключение кафедры.

Облачный Playwright работает в отдельном сеансе. Подключение к личному браузеру пользователя и его учётной записи СДО не установлено. Полученные учебные файлы позволяют продолжать работу без входа в СДО.
