(function () {
    let logoDataUriPromise = null;
    const monochromeLogoCache = new Map();
    const textEncoder = new TextEncoder();
    const isBatchContext = new URLSearchParams(window.location.search).get('officeExportContext') === 'batch';
    const currentHtmlFile = decodeURIComponent(window.location.pathname.split('/').pop() || '');
    const readyPromise = new Promise(function (resolve) {
        window.__officeFormsReadyResolve = resolve;
    });

    const PDF_PAGE_SIZE_IN = {
        a4: { width: 8.2677, height: 11.6929 },
        letter: { width: 8.5, height: 11 }
    };

    const OFFICE_FORM_FILES = [
        '3 COVID-Screening-Form.html',
        '4 Notice of Privacy Practice.html',
        '5 CT Scan Consent.html',
        'All on 4 Care & Maintainance Form.html',
        'BMP Information Packet.html',
        'BoneGraft.html',
        'CFID-Impression-Form.html',
        'Confidential Health History.html',
        'Consent - All on 4 Gum Disease.html',
        'Consent - Bisphosphonates Medications.html',
        'Consent - Block Graft.html',
        'Consent - Dental Implants.html',
        'Consent - Extraction.html',
        'Consent - Gingival Graft.html',
        'Consent - Graft of Max Sinus.html',
        'Consent - Guided Tissue.html',
        'Consent - IV Sedation.html',
        'Consent Sinus Lift + Implant.html',
        'Consent for Care of Dental Implants and Implant Restoration.html',
        'Consent for Frenectomy.html',
        'Consent- Single Implants.html',
        'Consent- SinusLift BoneGraft.html',
        'Exam Denial Consent.html',
        'Health History.html',
        'Upper & Lower Hybrid Impression Form.html',
        'INFORMED CONSENT DISCUSSION FOR LASER ASSISTED NEW ATTACHMENT PROCEDURE.html',
        'Immediate Implant Placement.html',
        'Implant Crown & Hybrid Denture Consent Form.html',
        'Implant Surgery Post Op instructions.html',
        'Implants Sticker.html',
        'Lab Form.html',
        'Informed Consent for Laser Dental Procedures.html',
        'Informed Consent.html',
        'Post Operative Care of Hybrid and Maintenance.html',
        'Pre-Post-Sedation Instructions.html',
        'PreConsultation Forms.html',
        'Smokingdocs.html',
        'To Operate or other procedure.html',
        'UPPER HYBRID PROPHYLAXIS CHECKLIST.html',
        'What to Eat After Implant Surgery.html',
        'allon4checklist.html',
        'biopsy_consent_form.html',
        'case log for justin.html',
        'consent for sinus grafting.html',
        'dental release form.html',
        'med_clearance_dental.html',
        'patient_registration.html',
        'teeth whitening.html'
    ];
    const OFFICE_FORM_FILE_SET = new Set(OFFICE_FORM_FILES);
    const DENSE_FORM_FILES = new Set([
        '3 COVID-Screening-Form.html',
        'Implants Sticker.html',
        'UPPER HYBRID PROPHYLAXIS CHECKLIST.html',
        'allon4checklist.html',
        'case log for justin.html',
        'med_clearance_dental.html'
    ]);
    const XDENSE_FORM_FILES = new Set([
        'CFID-Impression-Form.html',
        'Confidential Health History.html',
        'Consent - Bisphosphonates Medications.html',
        'Health History.html',
        'patient_registration.html',
        'PreConsultation Forms.html'
    ]);

    const CRC_TABLE = (function buildCrcTable() {
        const table = new Uint32Array(256);

        for (let index = 0; index < 256; index += 1) {
            let value = index;
            for (let bit = 0; bit < 8; bit += 1) {
                if ((value & 1) === 1) {
                    value = 0xedb88320 ^ (value >>> 1);
                } else {
                    value >>>= 1;
                }
            }
            table[index] = value >>> 0;
        }

        return table;
    }());

    function resolveReady() {
        if (typeof window.__officeFormsReadyResolve === 'function') {
            window.__officeFormsReadyResolve();
            window.__officeFormsReadyResolve = null;
        }
    }

    function normalizeWhitespace(text) {
        return String(text || '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function sanitizeFilename(text, fallback) {
        const cleaned = normalizeWhitespace(text)
            .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '')
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '')
            .toLowerCase();

        return cleaned || fallback;
    }

    function readDocumentTitle() {
        const visibleTitle = document.querySelector('.title-band .form-title, .premium-product-band .form-title');
        if (visibleTitle) {
            return normalizeWhitespace(visibleTitle.textContent);
        }

        return normalizeWhitespace(document.title).replace(/\s*\|\s*All-On-8 Robust\s*$/i, '');
    }

    function filenameFromTitle() {
        return sanitizeFilename(readDocumentTitle(), 'office-form');
    }

    function getPdfFileName() {
        return filenameFromTitle() + '.pdf';
    }

    function getPdfRoot() {
        return document.getElementById('pdf-content') || document.querySelector('.document');
    }

    function createLogoFallback() {
        const fallback = document.createElement('div');
        fallback.className = 'office-logo-fallback';
        fallback.innerHTML = [
            '<span class="office-logo-fallback-eyebrow">Center For</span>',
            '<span class="office-logo-fallback-title">Implant Dentistry</span>'
        ].join('');
        return fallback;
    }

    function ensureLogoDataUri() {
        if (window.OFFICE_FORMS_LOGO_DATA_URI) {
            return Promise.resolve(window.OFFICE_FORMS_LOGO_DATA_URI);
        }

        if (logoDataUriPromise) {
            return logoDataUriPromise;
        }

        logoDataUriPromise = new Promise(function (resolve) {
            const script = document.createElement('script');
            script.src = 'office-logo-data.js';
            script.async = true;
            script.dataset.officeLogoData = 'true';
            script.onload = function () {
                resolve(window.OFFICE_FORMS_LOGO_DATA_URI || '');
            };
            script.onerror = function () {
                resolve('');
            };
            document.head.appendChild(script);
        });

        return logoDataUriPromise;
    }

    function normalizeMasthead() {
        const masthead = document.querySelector('.masthead');
        if (!masthead) {
            return;
        }

        masthead.classList.add('office-print-avoid');

        const logo = masthead.querySelector('.brand-logo img, img[src$="logo.svg"], img[src*="logo"]');
        if (logo) {
            logo.setAttribute('alt', logo.getAttribute('alt') || 'Center For Implant Dentistry');
        }
    }

    function hydrateLogoAsset() {
        return ensureLogoDataUri().then(function (logoDataUri) {
            if (!logoDataUri) {
                return;
            }

            document.querySelectorAll('.masthead .brand-logo img, .masthead img[src$="logo.svg"], .masthead img[src*="logo"]').forEach(function (image) {
                const currentSrc = image.getAttribute('src') || '';

                if (currentSrc === logoDataUri) {
                    return;
                }

                image.setAttribute('src', logoDataUri);
                image.setAttribute('decoding', 'async');
                image.setAttribute('loading', 'eager');
            });
        });
    }

    function applyDocumentClasses() {
        const formSlug = sanitizeFilename(
            String(currentHtmlFile || '').replace(/\.html$/i, ''),
            'office-form'
        );
        const root = getPdfRoot();

        document.body.classList.add('office-form', 'office-form-' + formSlug);

        if (root) {
            root.classList.add('office-form-' + formSlug);
        }

        if (OFFICE_FORM_FILE_SET.has(currentHtmlFile)) {
            document.body.classList.add('office-form-dense');
        }

        if (DENSE_FORM_FILES.has(currentHtmlFile) || XDENSE_FORM_FILES.has(currentHtmlFile)) {
            document.body.classList.add('office-form-xdense');
        }

        if (document.querySelector('.premium-product-band')) {
            document.body.classList.add('office-form-has-premium');
        }
    }

    function inlinePrimaryTitleBand() {
        const masthead = document.querySelector('.masthead');
        const titleBand = document.querySelector('.title-band');

        if (!masthead || !titleBand || titleBand.classList.contains('office-inline-title')) {
            return;
        }

        if (document.querySelector('.premium-product-band')) {
            return;
        }

        const mastheadRight = masthead.querySelector('.masthead-right');
        masthead.classList.add('has-inline-title');
        titleBand.classList.add('office-inline-title', 'office-print-avoid');

        if (mastheadRight && mastheadRight.parentNode === masthead) {
            masthead.insertBefore(titleBand, mastheadRight);
            return;
        }

        masthead.appendChild(titleBand);
    }

    function normalizeManualPageBreaks() {
        document.querySelectorAll('.page-break').forEach(function (node) {
            node.classList.add('office-manual-page-break');
        });
    }

    function markAvoidBreaks() {
        const selectors = [
            '.premium-product-band',
            '.title-band',
            '.hhs-header',
            '.patient-header-row',
            '.signature-section',
            '.sig-section',
            '.patient-meta-band',
            '.important-box',
            '.consent-callout',
            '.notes-area',
            '.delivery-note',
            '.doc-footer'
        ].join(', ');

        document.querySelectorAll(selectors).forEach(function (node) {
            node.classList.add('office-print-avoid');
        });
    }

    function isVisibleForPagination(node) {
        if (!node || !node.getBoundingClientRect) {
            return false;
        }

        const style = window.getComputedStyle(node);
        if (style.display === 'none' || style.visibility === 'hidden') {
            return false;
        }

        const rect = node.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
    }

    function resolvePdfMargins(margin) {
        if (Array.isArray(margin)) {
            if (margin.length === 2) {
                return {
                    top: margin[0],
                    right: margin[1],
                    bottom: margin[0],
                    left: margin[1]
                };
            }

            if (margin.length === 4) {
                return {
                    top: margin[0],
                    right: margin[1],
                    bottom: margin[2],
                    left: margin[3]
                };
            }
        }

        const uniform = typeof margin === 'number' ? margin : 0;
        return {
            top: uniform,
            right: uniform,
            bottom: uniform,
            left: uniform
        };
    }

    function resolvePdfPageSize(pdfOptions) {
        const jsPdfOptions = (pdfOptions && pdfOptions.jsPDF) || {};
        const format = String(jsPdfOptions.format || 'a4').toLowerCase();
        const orientation = String(jsPdfOptions.orientation || 'portrait').toLowerCase();
        const baseSize = PDF_PAGE_SIZE_IN[format] || PDF_PAGE_SIZE_IN.a4;
        const isLandscape = orientation === 'landscape';

        return {
            width: isLandscape ? baseSize.height : baseSize.width,
            height: isLandscape ? baseSize.width : baseSize.height
        };
    }

    function getBreakTarget(node) {
        let target = node;
        const previousElement = node.previousElementSibling;

        if (
            previousElement &&
            (
                previousElement.classList.contains('part-header') ||
                previousElement.matches('h3')
            )
        ) {
            target = previousElement;
        }

        return target;
    }

    function preparePagination(root, pdfOptions) {
        const cleanupNodes = [];
        const margins = resolvePdfMargins(pdfOptions.margin);
        const pageSize = resolvePdfPageSize(pdfOptions);
        const printableWidthIn = pageSize.width - margins.left - margins.right;
        const printableHeightIn = pageSize.height - margins.top - margins.bottom;
        const rootRect = root.getBoundingClientRect();
        const contentWidthPx = root.scrollWidth || rootRect.width;
        const contentHeightPx = root.scrollHeight || rootRect.height;
        const pageHeightPx = printableWidthIn > 0
            ? (contentWidthPx * printableHeightIn) / printableWidthIn
            : 0;

        if (!pageHeightPx || !Number.isFinite(pageHeightPx)) {
            return function restorePagination() {
                return undefined;
            };
        }

        if (contentHeightPx <= pageHeightPx * 1.02) {
            return function restorePagination() {
                return undefined;
            };
        }

        const estimatedPages = Math.max(1, Math.ceil((contentHeightPx + 1) / pageHeightPx));

        if (estimatedPages <= 2) {
            return function restorePagination() {
                return undefined;
            };
        }

        document.documentElement.style.setProperty('--office-export-width', printableWidthIn.toFixed(4) + 'in');
        cleanupNodes.push({ node: document.documentElement, className: '--office-export-width' });

        root.querySelectorAll('.office-print-break-before').forEach(function (node) {
            node.classList.remove('office-print-break-before');
        });

        const candidateSelector = [
            '.part-header',
            'h3',
            '.notes-area',
            '.delivery-note',
            '.signature-section',
            '.sig-section',
            '.important-box',
            '.consent-callout'
        ].join(', ');

        for (let pass = 0; pass < 6; pass += 1) {
            const pageTop = root.getBoundingClientRect().top;
            let changed = false;

            root.querySelectorAll(candidateSelector).forEach(function (node) {
                if (changed || !isVisibleForPagination(node)) {
                    return;
                }

                const rect = node.getBoundingClientRect();
                const top = rect.top - pageTop;
                const bottom = rect.bottom - pageTop;
                const height = rect.height;

                if (height >= pageHeightPx * 0.9 || top <= 8) {
                    return;
                }

                const currentPage = Math.floor(top / pageHeightPx);
                const currentPageTop = currentPage * pageHeightPx;
                const currentPageBottom = currentPageTop + pageHeightPx;
                const remainingSpace = currentPageBottom - top;

                if (bottom <= currentPageBottom - 4 || top <= currentPageTop + 8) {
                    return;
                }

                const isHeading = node.matches('.part-header, h3');
                const isCompactBlock = isHeading || node.matches('.notes-area, .delivery-note, .signature-section, .sig-section');
                const headingThreshold = estimatedPages > 2
                    ? Math.min(pageHeightPx * 0.12, 84)
                    : Math.min(pageHeightPx * 0.08, 60);
                const blockThreshold = estimatedPages > 2
                    ? Math.min(pageHeightPx * 0.18, 148)
                    : Math.min(pageHeightPx * 0.12, 98);

                if (isHeading && remainingSpace > headingThreshold) {
                    return;
                }

                if (!isHeading && (!isCompactBlock || remainingSpace > Math.max(height + 10, blockThreshold))) {
                    return;
                }

                const breakTarget = getBreakTarget(node);
                if (!breakTarget.classList.contains('office-print-break-before')) {
                    breakTarget.classList.add('office-print-break-before');
                    cleanupNodes.push({ node: breakTarget, className: 'office-print-break-before' });
                    changed = true;
                }
            });

            if (!changed) {
                break;
            }
        }

        return function restorePagination() {
            cleanupNodes.forEach(function (entry) {
                if (entry.className === '--office-export-width') {
                    entry.node.style.removeProperty('--office-export-width');
                    return;
                }

                entry.node.classList.remove(entry.className);
            });
        };
    }

    function buildPdfOptions(filename) {
        return {
            margin: [0.14, 0.14, 0.18, 0.14],
            filename: filename + '.pdf',
            image: { type: 'jpeg', quality: 1 },
            html2canvas: {
                scale: Math.min(2.6, Math.max(2.3, (window.devicePixelRatio || 1) * 1.8)),
                useCORS: true,
                backgroundColor: '#ffffff',
                letterRendering: true,
                logging: false,
                removeContainer: true,
                scrollX: 0,
                scrollY: 0
            },
            jsPDF: {
                unit: 'in',
                format: 'a4',
                orientation: 'portrait',
                compress: true
            },
            pagebreak: {
                mode: ['css'],
                avoid: [
                    '.office-print-avoid',
                    '.part-header',
                    'h3',
                    '.signature-section',
                    '.sig-section',
                    '.notes-area',
                    '.delivery-note',
                    'p',
                    'li',
                    'tr',
                    '.field-row',
                    '.form-row',
                    '.question-block',
                    '.checkbox-item'
                ]
            }
        };
    }

    function prepareExportLayout(pdfOptions) {
        const margins = resolvePdfMargins(pdfOptions.margin);
        const pageSize = resolvePdfPageSize(pdfOptions);
        const printableWidthIn = Math.max(1, pageSize.width - margins.left - margins.right);

        document.documentElement.style.setProperty('--office-export-width', printableWidthIn.toFixed(4) + 'in');

        return function restoreExportLayout() {
            document.documentElement.style.removeProperty('--office-export-width');
        };
    }

    function waitForImageReady(image) {
        if (typeof image.decode === 'function') {
            return image.decode().catch(function () {
                return undefined;
            });
        }

        return new Promise(function (resolve) {
            if (image.complete) {
                resolve();
                return;
            }

            function finish() {
                image.removeEventListener('load', finish);
                image.removeEventListener('error', finish);
                resolve();
            }

            image.addEventListener('load', finish);
            image.addEventListener('error', finish);
        });
    }

    function buildMonochromeLogoDataUri(src) {
        if (!src) {
            return Promise.resolve('');
        }

        if (monochromeLogoCache.has(src)) {
            return Promise.resolve(monochromeLogoCache.get(src));
        }

        return new Promise(function (resolve) {
            const image = new Image();
            image.crossOrigin = 'anonymous';

            image.onload = function () {
                try {
                    const width = image.naturalWidth || image.width;
                    const height = image.naturalHeight || image.height;

                    if (!width || !height) {
                        monochromeLogoCache.set(src, '');
                        resolve('');
                        return;
                    }

                    const canvas = document.createElement('canvas');
                    const upscale = Math.max(1, width < 1600 ? 2 : 1.5);
                    canvas.width = Math.round(width * upscale);
                    canvas.height = Math.round(height * upscale);

                    const context = canvas.getContext('2d');
                    if (!context) {
                        monochromeLogoCache.set(src, '');
                        resolve('');
                        return;
                    }

                    context.imageSmoothingEnabled = true;
                    context.imageSmoothingQuality = 'high';
                    context.drawImage(image, 0, 0, canvas.width, canvas.height);

                    const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
                    const pixels = imageData.data;

                    for (let index = 0; index < pixels.length; index += 4) {
                        const alpha = pixels[index + 3];
                        if (!alpha) {
                            continue;
                        }

                        const luminance = Math.round(
                            (pixels[index] * 0.2126) +
                            (pixels[index + 1] * 0.7152) +
                            (pixels[index + 2] * 0.0722)
                        );

                        const adjusted = luminance > 245
                            ? 255
                            : Math.max(0, Math.min(255, Math.round(luminance * 0.35)));

                        pixels[index] = adjusted;
                        pixels[index + 1] = adjusted;
                        pixels[index + 2] = adjusted;
                    }

                    context.putImageData(imageData, 0, 0);

                    const dataUri = canvas.toDataURL('image/png');
                    monochromeLogoCache.set(src, dataUri);
                    resolve(dataUri);
                } catch (error) {
                    monochromeLogoCache.set(src, '');
                    resolve('');
                }
            };

            image.onerror = function () {
                monochromeLogoCache.set(src, '');
                resolve('');
            };

            image.src = src;
        });
    }

    function prepareExportResources(root) {
        const restorers = [];
        const pending = [];
        const inlineLogo = window.OFFICE_FORMS_LOGO_DATA_URI || '';
        const logoSelector = '.masthead .brand-logo img, .masthead img[src$="logo.svg"], .masthead img[src*="logo"]';

        root.querySelectorAll(logoSelector).forEach(function (image) {
            const parent = image.parentNode;
            if (!parent) {
                return;
            }

            const originalSrc = image.getAttribute('src') || '';
            const preferredSrc = inlineLogo || originalSrc;

            pending.push(
                buildMonochromeLogoDataUri(preferredSrc).then(function (monochromeSrc) {
                    if (monochromeSrc) {
                        image.setAttribute('src', monochromeSrc);
                        return waitForImageReady(image).then(function () {
                            restorers.push(function () {
                                image.setAttribute('src', originalSrc);
                            });
                        });
                    }

                    if (inlineLogo) {
                        image.setAttribute('src', inlineLogo);
                        return waitForImageReady(image).then(function () {
                            restorers.push(function () {
                                image.setAttribute('src', originalSrc);
                            });
                        });
                    }

                    const fallback = createLogoFallback();
                    const nextSibling = image.nextSibling;
                    parent.replaceChild(fallback, image);

                    restorers.push(function () {
                        if (nextSibling && nextSibling.parentNode === parent) {
                            parent.insertBefore(image, nextSibling);
                        } else {
                            parent.appendChild(image);
                        }
                        fallback.remove();
                    });
                })
            );
        });

        return Promise.all(pending).then(function () {
            return function restoreResources() {
                while (restorers.length) {
                    restorers.pop()();
                }
            };
        });
    }

    function extractRenderedImageHeight(commands) {
        if (!Array.isArray(commands)) {
            return null;
        }

        const matrixCommand = commands.find(function (command) {
            return typeof command === 'string' && /\scm$/.test(command);
        });

        if (!matrixCommand) {
            return null;
        }

        const values = matrixCommand
            .replace(/\scm$/, '')
            .trim()
            .split(/\s+/)
            .map(function (value) {
                return Number(value);
            });

        if (values.length < 6 || values.some(function (value) { return !Number.isFinite(value); })) {
            return null;
        }

        return Math.abs(values[3]);
    }

    function pruneArtifactPages(pdf) {
        if (!pdf || typeof pdf.deletePage !== 'function' || !pdf.internal || !Array.isArray(pdf.internal.pages)) {
            return;
        }

        for (let pageIndex = pdf.internal.pages.length - 1; pageIndex >= 1; pageIndex -= 1) {
            const commands = pdf.internal.pages[pageIndex];
            if (!Array.isArray(commands) || commands.length > 6) {
                continue;
            }

            const renderedHeight = extractRenderedImageHeight(commands);
            if (renderedHeight !== null && renderedHeight < 40) {
                pdf.deletePage(pageIndex);
            }
        }
    }

    function getLayoutMetrics() {
        const root = getPdfRoot();
        if (!root) {
            return null;
        }

        const pdfOptions = buildPdfOptions(filenameFromTitle());
        const margins = resolvePdfMargins(pdfOptions.margin);
        const pageSize = resolvePdfPageSize(pdfOptions);
        const rect = root.getBoundingClientRect();
        const contentWidthPx = root.scrollWidth || rect.width || 1;
        const contentHeightPx = root.scrollHeight || rect.height || 0;
        const printableWidthIn = pageSize.width - margins.left - margins.right;
        const printableHeightIn = pageSize.height - margins.top - margins.bottom;
        const printablePageHeightPx = printableWidthIn > 0
            ? (contentWidthPx * printableHeightIn) / printableWidthIn
            : 0;

        return {
            file: currentHtmlFile,
            title: readDocumentTitle(),
            widthPx: contentWidthPx,
            heightPx: contentHeightPx,
            printablePageHeightPx: printablePageHeightPx,
            estimatedPages: printablePageHeightPx
                ? Math.max(1, Math.ceil((contentHeightPx + 1) / printablePageHeightPx))
                : 1
        };
    }

    async function renderPdf(root, filename) {
        if (!root || typeof window.html2pdf !== 'function') {
            throw new Error('html2pdf is not available');
        }

        const pdfOptions = buildPdfOptions(filename);
        await ensureLogoDataUri();
        const restoreResources = await prepareExportResources(root);
        const restoreLayout = prepareExportLayout(pdfOptions);
        document.body.classList.add('office-forms-exporting');

        try {
            return await html2pdf()
                .set(pdfOptions)
                .from(root)
                .toPdf()
                .get('pdf');
        } finally {
            restoreLayout();
            restoreResources();
            document.body.classList.remove('office-forms-exporting');
        }
    }

    async function measurePdfPageCount() {
        const root = getPdfRoot();
        const filename = filenameFromTitle();
        const pdf = await renderPdf(root, filename);
        pruneArtifactPages(pdf);

        return {
            file: currentHtmlFile,
            fileName: filename + '.pdf',
            title: readDocumentTitle(),
            pageCount: pdf.internal.getNumberOfPages()
        };
    }

    async function createPdfBlob(root, filename) {
        const pdf = await renderPdf(root, filename);
        pruneArtifactPages(pdf);
        return pdf.output('blob');
    }

    async function exportPdfBlob() {
        const root = getPdfRoot();
        const filename = filenameFromTitle();
        const blob = await createPdfBlob(root, filename);

        return {
            blob: blob,
            fileName: filename + '.pdf',
            title: readDocumentTitle()
        };
    }

    function saveBlob(blob, filename) {
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = filename;
        link.rel = 'noopener';
        document.body.appendChild(link);
        link.click();
        link.remove();

        window.setTimeout(function () {
            URL.revokeObjectURL(downloadUrl);
        }, 2000);
    }

    function setActionState(buttons, state) {
        if (!buttons || !buttons.pdfButton || !buttons.allButton) {
            return;
        }

        const pdfLabel = state && state.pdfLabel ? state.pdfLabel : 'Download PDF';
        const allLabel = state && state.allLabel ? state.allLabel : 'Download All';
        const isBusy = Boolean(state && state.busy);

        buttons.pdfButton.disabled = isBusy;
        buttons.allButton.disabled = isBusy;
        buttons.pdfButton.textContent = pdfLabel;
        buttons.allButton.textContent = allLabel;
    }

    function normalizeActionButtons() {
        const previousGroup = document.querySelector('.office-export-actions');
        const previousPdfButton = document.getElementById('downloadPdfBtn');
        const previousAllButton = document.getElementById('downloadAllBtn');

        if (previousGroup) {
            previousGroup.remove();
        }

        if (previousPdfButton) {
            previousPdfButton.remove();
        }

        if (previousAllButton) {
            previousAllButton.remove();
        }

        const group = document.createElement('div');
        group.className = 'office-export-actions';

        const pdfButton = document.createElement('button');
        pdfButton.id = 'downloadPdfBtn';
        pdfButton.className = 'fab';
        pdfButton.type = 'button';
        pdfButton.textContent = 'Download PDF';

        const allButton = document.createElement('button');
        allButton.id = 'downloadAllBtn';
        allButton.className = 'fab';
        allButton.type = 'button';
        allButton.textContent = 'Download All';

        group.appendChild(pdfButton);
        group.appendChild(allButton);
        document.body.appendChild(group);

        return {
            group: group,
            pdfButton: pdfButton,
            allButton: allButton
        };
    }

    function crc32(bytes) {
        let crc = 0xffffffff;

        for (let index = 0; index < bytes.length; index += 1) {
            crc = CRC_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
        }

        return (crc ^ 0xffffffff) >>> 0;
    }

    function writeUint16(view, offset, value) {
        view[offset] = value & 0xff;
        view[offset + 1] = (value >>> 8) & 0xff;
    }

    function writeUint32(view, offset, value) {
        view[offset] = value & 0xff;
        view[offset + 1] = (value >>> 8) & 0xff;
        view[offset + 2] = (value >>> 16) & 0xff;
        view[offset + 3] = (value >>> 24) & 0xff;
    }

    function toDosDateTime(date) {
        const normalized = new Date(date);
        const year = Math.max(1980, normalized.getFullYear());

        return {
            time: ((normalized.getHours() & 0x1f) << 11) |
                ((normalized.getMinutes() & 0x3f) << 5) |
                Math.floor(normalized.getSeconds() / 2),
            date: (((year - 1980) & 0x7f) << 9) |
                (((normalized.getMonth() + 1) & 0x0f) << 5) |
                (normalized.getDate() & 0x1f)
        };
    }

    function createZipState() {
        return {
            localParts: [],
            centralParts: [],
            localOffset: 0,
            entryCount: 0,
            timestamp: toDosDateTime(new Date())
        };
    }

    async function appendZipEntry(zipState, fileName, blob) {
        const nameBytes = textEncoder.encode(fileName);
        const fileBytes = new Uint8Array(await blob.arrayBuffer());
        const checksum = crc32(fileBytes);

        const localHeader = new Uint8Array(30 + nameBytes.length);
        writeUint32(localHeader, 0, 0x04034b50);
        writeUint16(localHeader, 4, 20);
        writeUint16(localHeader, 6, 0);
        writeUint16(localHeader, 8, 0);
        writeUint16(localHeader, 10, zipState.timestamp.time);
        writeUint16(localHeader, 12, zipState.timestamp.date);
        writeUint32(localHeader, 14, checksum);
        writeUint32(localHeader, 18, fileBytes.length);
        writeUint32(localHeader, 22, fileBytes.length);
        writeUint16(localHeader, 26, nameBytes.length);
        writeUint16(localHeader, 28, 0);
        localHeader.set(nameBytes, 30);

        zipState.localParts.push(localHeader, fileBytes);

        const centralHeader = new Uint8Array(46 + nameBytes.length);
        writeUint32(centralHeader, 0, 0x02014b50);
        writeUint16(centralHeader, 4, 20);
        writeUint16(centralHeader, 6, 20);
        writeUint16(centralHeader, 8, 0);
        writeUint16(centralHeader, 10, 0);
        writeUint16(centralHeader, 12, zipState.timestamp.time);
        writeUint16(centralHeader, 14, zipState.timestamp.date);
        writeUint32(centralHeader, 16, checksum);
        writeUint32(centralHeader, 20, fileBytes.length);
        writeUint32(centralHeader, 24, fileBytes.length);
        writeUint16(centralHeader, 28, nameBytes.length);
        writeUint16(centralHeader, 30, 0);
        writeUint16(centralHeader, 32, 0);
        writeUint16(centralHeader, 34, 0);
        writeUint16(centralHeader, 36, 0);
        writeUint32(centralHeader, 38, 0);
        writeUint32(centralHeader, 42, zipState.localOffset);
        centralHeader.set(nameBytes, 46);

        zipState.centralParts.push(centralHeader);
        zipState.localOffset += localHeader.length + fileBytes.length;
        zipState.entryCount += 1;
    }

    async function cloneBlob(blob) {
        if (!blob) {
            return new Blob([]);
        }

        return new Blob([await blob.arrayBuffer()], {
            type: blob.type || 'application/octet-stream'
        });
    }

    function finalizeZipBlob(zipState) {
        const centralSize = zipState.centralParts.reduce(function (total, part) {
            return total + part.length;
        }, 0);

        const endRecord = new Uint8Array(22);
        writeUint32(endRecord, 0, 0x06054b50);
        writeUint16(endRecord, 4, 0);
        writeUint16(endRecord, 6, 0);
        writeUint16(endRecord, 8, zipState.entryCount);
        writeUint16(endRecord, 10, zipState.entryCount);
        writeUint32(endRecord, 12, centralSize);
        writeUint32(endRecord, 16, zipState.localOffset);
        writeUint16(endRecord, 20, 0);

        return new Blob(zipState.localParts.concat(zipState.centralParts).concat(endRecord), {
            type: 'application/zip'
        });
    }

    async function createZipBlob(entries) {
        const zipState = createZipState();

        for (let index = 0; index < entries.length; index += 1) {
            await appendZipEntry(zipState, entries[index].fileName, entries[index].blob);
        }

        return finalizeZipBlob(zipState);
    }

    function dedupeFileName(fileName, usedNames) {
        const extensionIndex = fileName.lastIndexOf('.');
        const baseName = extensionIndex >= 0 ? fileName.slice(0, extensionIndex) : fileName;
        const extension = extensionIndex >= 0 ? fileName.slice(extensionIndex) : '';
        let candidate = fileName;
        let counter = 2;

        while (usedNames.has(candidate.toLowerCase())) {
            candidate = baseName + '-' + counter + extension;
            counter += 1;
        }

        usedNames.add(candidate.toLowerCase());
        return candidate;
    }

    function appendQuery(url, key, value) {
        const joiner = url.indexOf('?') === -1 ? '?' : '&';
        return url + joiner + encodeURIComponent(key) + '=' + encodeURIComponent(value);
    }

    function waitForApi(frameWindow, timeoutMs, sourceFile) {
        const timeout = typeof timeoutMs === 'number' ? timeoutMs : 15000;
        const startedAt = Date.now();

        return new Promise(function (resolve, reject) {
            function poll() {
                if (
                    frameWindow &&
                    frameWindow.OfficeFormsPrint &&
                    typeof frameWindow.OfficeFormsPrint.exportPdfBlob === 'function'
                ) {
                    resolve(frameWindow.OfficeFormsPrint);
                    return;
                }

                if (Date.now() - startedAt >= timeout) {
                    reject(new Error('Timed out waiting for embedded form export API for ' + sourceFile));
                    return;
                }

                window.setTimeout(poll, 120);
            }

            poll();
        });
    }

    async function loadBatchFrame(sourceFile) {
        const iframe = document.createElement('iframe');
        iframe.className = 'office-export-frame';
        iframe.src = appendQuery(encodeURI(sourceFile), 'officeExportContext', 'batch');

        const loadPromise = new Promise(function (resolve, reject) {
            iframe.onload = function () {
                resolve();
            };
            iframe.onerror = function () {
                reject(new Error('Failed to load ' + sourceFile));
            };
        });

        document.body.appendChild(iframe);
        await loadPromise;

        const api = await waitForApi(iframe.contentWindow, 20000, sourceFile);
        if (typeof api.whenReady === 'function') {
            await api.whenReady();
        }

        return {
            iframe: iframe,
            api: api
        };
    }

    async function handleSingleDownload(buttons) {
        setActionState(buttons, {
            busy: true,
            pdfLabel: 'Preparing PDF',
            allLabel: 'Download All'
        });

        try {
            const result = await exportPdfBlob();
            saveBlob(result.blob, result.fileName);
        } catch (error) {
            console.error('PDF export failed', error);
        } finally {
            setActionState(buttons);
        }
    }

    async function buildAllFormsZip(progressCallback) {
        const usedNames = new Set();
        const zipState = createZipState();

        for (let index = 0; index < OFFICE_FORM_FILES.length; index += 1) {
            const sourceFile = OFFICE_FORM_FILES[index];

            if (typeof progressCallback === 'function') {
                progressCallback({
                    phase: 'prepare',
                    index: index + 1,
                    total: OFFICE_FORM_FILES.length,
                    sourceFile: sourceFile
                });
            }

            if (sourceFile === currentHtmlFile) {
                const currentResult = await exportPdfBlob();
                await appendZipEntry(
                    zipState,
                    dedupeFileName(currentResult.fileName, usedNames),
                    currentResult.blob
                );
            } else {
                const frameHandle = await loadBatchFrame(sourceFile);

                try {
                    const frameResult = await frameHandle.api.exportPdfBlob();
                    const copiedBlob = await cloneBlob(frameResult.blob);

                    await appendZipEntry(
                        zipState,
                        dedupeFileName(frameResult.fileName, usedNames),
                        copiedBlob
                    );
                } finally {
                    frameHandle.iframe.remove();
                }
            }
        }

        if (typeof progressCallback === 'function') {
            progressCallback({
                phase: 'package',
                index: OFFICE_FORM_FILES.length,
                total: OFFICE_FORM_FILES.length,
                sourceFile: null
            });
        }

        return {
            blob: finalizeZipBlob(zipState),
            fileName: 'office-forms.zip'
        };
    }

    async function handleBatchDownload(buttons) {
        try {
            const zipResult = await buildAllFormsZip(function (progress) {
                if (progress.phase === 'package') {
                    setActionState(buttons, {
                        busy: true,
                        pdfLabel: 'Download PDF',
                        allLabel: 'Packaging ZIP'
                    });
                    return;
                }

                setActionState(buttons, {
                    busy: true,
                    pdfLabel: 'Download PDF',
                    allLabel: 'Preparing ' + progress.index + '/' + progress.total
                });
            });

            saveBlob(zipResult.blob, zipResult.fileName);
        } catch (error) {
            console.error('Batch PDF export failed', error);
        } finally {
            setActionState(buttons);
        }
    }

    function init() {
        if (!document.body || document.body.dataset.officeFormsEnhanced === 'true') {
            resolveReady();
            return;
        }

        document.body.dataset.officeFormsEnhanced = 'true';
        applyDocumentClasses();
        normalizeMasthead();
        inlinePrimaryTitleBand();
        normalizeManualPageBreaks();
        markAvoidBreaks();
        const logoReady = hydrateLogoAsset();

        if (!isBatchContext) {
            const buttons = normalizeActionButtons();
            buttons.pdfButton.addEventListener('click', function (event) {
                event.preventDefault();
                window.scrollTo(0, 0);
                handleSingleDownload(buttons);
            });
            buttons.allButton.addEventListener('click', function (event) {
                event.preventDefault();
                window.scrollTo(0, 0);
                handleBatchDownload(buttons);
            });
        }

        logoReady.finally(resolveReady);
    }

    window.OfficeFormsPrint = {
        whenReady: function () {
            return readyPromise;
        },
        exportPdfBlob: exportPdfBlob,
        buildAllFormsZip: buildAllFormsZip,
        getDocumentTitle: readDocumentTitle,
        getPdfFileName: getPdfFileName,
        getLayoutMetrics: getLayoutMetrics,
        measurePdfPageCount: measurePdfPageCount
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init, { once: true });
    } else {
        init();
    }
}());
