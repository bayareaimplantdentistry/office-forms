(function () {
    let logoDataUriPromise = null;
    let jsZipPromise = null;
    let batchExportRequestCounter = 0;
    let batchExportMessagingBound = false;
    const monochromeLogoCache = new Map();
    const BATCH_EXPORT_SCALE = 4;
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
        'INFORMED CONSENT DISCUSSION FOR LASER ASSISTED NEW ATTACHMENT PROCEDURE.html',
        'Immediate Implant Placement.html',
        'Implant Crown & Hybrid Denture Consent Form.html',
        'Implant Surgery Post Op instructions.html',
        'Implants Sticker.html',
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

    const PDF_PAGE_SIZE_IN = {
        a4: { width: 8.2677, height: 11.6929 },
        letter: { width: 8.5, height: 11 }
    };

    function createLogoFallback() {
        const fallback = document.createElement('div');
        fallback.className = 'office-logo-fallback';
        fallback.innerHTML = [
            '<span class="office-logo-fallback-eyebrow">Bay Area</span>',
            '<span class="office-logo-fallback-title">Implant Dentistry</span>'
        ].join('');
        return fallback;
    }

    function isBatchFrameMode() {
        let searchMode = '';

        try {
            searchMode = new window.URLSearchParams(window.location.search).get('office-export-mode') || '';
        } catch (error) {
            searchMode = '';
        }

        return searchMode === 'batch-frame' || window.self !== window.top;
    }

    function ensureLogoDataUri() {
        if (window.location.protocol !== 'file:') {
            return Promise.resolve('');
        }

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

    function normalizeWhitespace(text) {
        return String(text || '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function readDocumentTitle() {
        const visibleTitle = document.querySelector('.title-band .form-title');
        if (visibleTitle) {
            return normalizeWhitespace(visibleTitle.textContent);
        }

        return normalizeWhitespace(document.title).replace(/\s*\|\s*All-On-8 Robust\s*$/i, '');
    }

    function filenameFromTitle() {
        const cleanedTitle = readDocumentTitle()
            .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '')
            .trim()
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-')
            .toLowerCase();

        return cleanedTitle || 'office-form';
    }

    function getPdfFilename() {
        return filenameFromTitle() + '.pdf';
    }

    function formClassFromTitle() {
        const className = readDocumentTitle()
            .replace(/&/g, ' and ')
            .replace(/[^a-z0-9]+/gi, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '')
            .toLowerCase();

        return className || 'office-form';
    }

    function markFormIdentity() {
        const root = document.getElementById('pdf-content') || document.querySelector('.document');
        const formClass = formClassFromTitle();

        document.body.classList.add('office-form', 'office-form--' + formClass);
        document.body.dataset.officeForm = formClass;

        if (root) {
            root.classList.add('office-form-document', 'office-form-document--' + formClass);
            root.dataset.officeForm = formClass;
        }
    }

    function getExportRoot() {
        return document.getElementById('pdf-content') || document.querySelector('.document');
    }

    function normalizeMasthead() {
        const masthead = document.querySelector('.masthead');
        if (!masthead) {
            return;
        }

        masthead.classList.add('office-print-avoid');

        const logo = masthead.querySelector('.brand-logo img, img[src$="logo.svg"], img[src*="logo"]');
        if (logo) {
            logo.setAttribute('alt', logo.getAttribute('alt') || 'Bay Area Implant Dentistry');
            logo.style.filter = 'grayscale(100%) contrast(136%) brightness(0.88)';
        }
    }

    function inlinePrimaryTitleBand() {
        const masthead = document.querySelector('.document > .masthead, .masthead');
        if (!masthead) {
            return;
        }

        const titleBand = masthead.nextElementSibling;
        if (!titleBand || !titleBand.classList.contains('title-band')) {
            return;
        }

        if (!titleBand.querySelector('.form-title') || titleBand.classList.contains('office-inline-title')) {
            return;
        }

        masthead.classList.add('has-inline-title');
        titleBand.classList.add('office-inline-title');
        masthead.appendChild(titleBand);
    }

    function removeLegacyArtifacts() {
        document.querySelectorAll('.masthead-right, .premium-product-band, .doc-footer, .page-break').forEach(function (node) {
            node.remove();
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
            '.delivery-note'
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
        const format = String(jsPdfOptions.format || 'letter').toLowerCase();
        const orientation = String(jsPdfOptions.orientation || 'portrait').toLowerCase();
        const baseSize = PDF_PAGE_SIZE_IN[format] || PDF_PAGE_SIZE_IN.letter;
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
        const pageHeightPx = printableWidthIn > 0
            ? (contentWidthPx * printableHeightIn) / printableWidthIn
            : 0;

        if (!pageHeightPx || !Number.isFinite(pageHeightPx)) {
            return function restorePagination() {
                return undefined;
            };
        }

        document.documentElement.style.setProperty('--office-export-width', printableWidthIn.toFixed(4) + 'in');
        cleanupNodes.push({
            node: document.documentElement,
            className: '--office-export-width'
        });

        root.querySelectorAll('.office-print-break-before').forEach(function (node) {
            node.classList.remove('office-print-break-before');
        });

        const candidateSelector = [
            '.part-header',
            'h3',
            '.notes-area',
            '.delivery-note'
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

                if (height >= pageHeightPx * 0.92 || top <= 8) {
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
                if (isHeading && node.closest('.signature-section, .sig-section')) {
                    return;
                }

                const isCompactBlock = isHeading || node.matches('.notes-area, .delivery-note');
                const headingThreshold = Math.min(pageHeightPx * 0.14, 96);
                const blockThreshold = Math.min(pageHeightPx * 0.24, 180);

                if (isHeading && remainingSpace > headingThreshold) {
                    return;
                }

                if (!isHeading && (!isCompactBlock || height >= pageHeightPx * 0.45 || remainingSpace > Math.max(height + 16, blockThreshold))) {
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

    function normalizeExportButtons() {
        const existingDownload = document.getElementById('downloadPdfBtn');
        const existingBatch = document.getElementById('downloadAllBtn');
        let toolbar = document.getElementById('officeExportToolbar');

        if (!toolbar) {
            toolbar = document.createElement('div');
            toolbar.id = 'officeExportToolbar';
            toolbar.className = 'office-export-toolbar';
            document.body.appendChild(toolbar);
        }

        if (existingDownload) {
            existingDownload.remove();
        }

        if (existingBatch) {
            existingBatch.remove();
        }

        const downloadButton = document.createElement('button');
        downloadButton.id = 'downloadPdfBtn';
        downloadButton.className = 'fab';
        downloadButton.type = 'button';
        downloadButton.textContent = 'Download PDF';

        const batchButton = document.createElement('button');
        batchButton.id = 'downloadAllBtn';
        batchButton.className = 'fab fab-secondary';
        batchButton.type = 'button';
        batchButton.textContent = 'Download All';

        toolbar.replaceChildren(downloadButton, batchButton);

        return {
            toolbar: toolbar,
            download: downloadButton,
            batch: batchButton
        };
    }

    function buildPdfOptions(filename, exportOverrides) {
        const overrides = exportOverrides || {};
        const html2canvasScale = typeof overrides.scale === 'number'
            ? overrides.scale
            : Math.min(3.4, Math.max(2.8, window.devicePixelRatio || 1));

        return {
            margin: [0.18, 0.18, 0.18, 0.18],
            filename: overrides.filename || (filename + '.pdf'),
            image: { type: 'jpeg', quality: 1 },
            html2canvas: {
                scale: html2canvasScale,
                useCORS: true,
                backgroundColor: '#ffffff',
                letterRendering: true,
                removeContainer: true,
                scrollX: 0,
                scrollY: 0
            },
            jsPDF: {
                unit: 'in',
                format: 'a4',
                orientation: 'portrait',
                compress: typeof overrides.compress === 'boolean' ? overrides.compress : true
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
                    '.delivery-note'
                ]
            }
        };
    }

    function snapshotToolbarLabels() {
        return Array.prototype.reduce.call(
            document.querySelectorAll('.office-export-toolbar .fab'),
            function (labels, button) {
                labels[button.id] = button.textContent;
                return labels;
            },
            {}
        );
    }

    function restoreToolbar(labels) {
        document.body.classList.remove('office-forms-exporting');

        document.querySelectorAll('.office-export-toolbar .fab').forEach(function (button) {
            button.disabled = false;
            if (labels && Object.prototype.hasOwnProperty.call(labels, button.id)) {
                button.textContent = labels[button.id];
            }
        });
    }

    function disableToolbar() {
        document.body.classList.add('office-forms-exporting');
        document.querySelectorAll('.office-export-toolbar .fab').forEach(function (button) {
            button.disabled = true;
        });
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

    function runPdfExport(exportOverrides, workerAction) {
        const root = getExportRoot();

        if (!root) {
            return Promise.reject(new Error('No printable form root found.'));
        }

        if (typeof window.html2pdf !== 'function') {
            return Promise.reject(new Error('html2pdf is not available.'));
        }

        const pdfOptions = buildPdfOptions(filenameFromTitle(), exportOverrides);
        window.scrollTo(0, 0);

        return ensureLogoDataUri()
            .then(function () {
                return prepareExportResources(root);
            })
            .then(function (restoreResources) {
                const restorePagination = preparePagination(root, pdfOptions);
                const worker = html2pdf()
                    .set(pdfOptions)
                    .from(root);

                return Promise.resolve()
                    .then(function () {
                        return workerAction(worker, pdfOptions);
                    })
                    .then(function (result) {
                        restorePagination();
                        restoreResources();
                        return result;
                    })
                    .catch(function (error) {
                        restorePagination();
                        restoreResources();
                        throw error;
                    });
            });
    }

    function exportCurrentPagePdfToArrayBuffer(exportOverrides) {
        return runPdfExport(exportOverrides, function (worker, pdfOptions) {
            return worker
                .toPdf()
                .get('pdf')
                .then(function (pdf) {
                    return {
                        data: pdf.output('arraybuffer'),
                        filename: pdfOptions.filename
                    };
                });
        });
    }

    function setupBatchExportMessaging() {
        if (batchExportMessagingBound) {
            return;
        }

        batchExportMessagingBound = true;

        window.addEventListener('message', function (event) {
            const payload = event.data || {};

            if (payload.officeFormsType !== 'office-forms-export-request' || !event.source) {
                return;
            }

            exportCurrentPagePdfToArrayBuffer(payload.exportOptions)
                .then(function (result) {
                    event.source.postMessage({
                        officeFormsType: 'office-forms-export-result',
                        requestId: payload.requestId,
                        filename: result.filename,
                        data: result.data
                    }, '*', [result.data]);
                })
                .catch(function (error) {
                    event.source.postMessage({
                        officeFormsType: 'office-forms-export-error',
                        requestId: payload.requestId,
                        message: error && error.message ? error.message : 'Form export failed.'
                    }, '*');
                });
        });
    }

    function normalizeManifestEntry(entry) {
        if (typeof entry === 'string') {
            return entry;
        }

        if (entry && typeof entry.path === 'string') {
            return entry.path;
        }

        return '';
    }

    function getFormManifest() {
        const source = Array.isArray(window.OFFICE_FORMS_MANIFEST) && window.OFFICE_FORMS_MANIFEST.length
            ? window.OFFICE_FORMS_MANIFEST
            : OFFICE_FORM_FILES;
        const seen = Object.create(null);

        return source
            .map(normalizeManifestEntry)
            .filter(function (path) {
                if (!path || seen[path]) {
                    return false;
                }

                seen[path] = true;
                return true;
            });
    }

    function ensureJsZip() {
        if (window.JSZip) {
            return Promise.resolve(window.JSZip);
        }

        if (jsZipPromise) {
            return jsZipPromise;
        }

        jsZipPromise = new Promise(function (resolve, reject) {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            script.async = true;
            script.dataset.officeFormsJszip = 'true';
            script.onload = function () {
                if (window.JSZip) {
                    resolve(window.JSZip);
                    return;
                }

                reject(new Error('JSZip did not load.'));
            };
            script.onerror = function () {
                reject(new Error('Unable to load JSZip.'));
            };
            document.head.appendChild(script);
        }).catch(function (error) {
            jsZipPromise = null;
            throw error;
        });

        return jsZipPromise;
    }

    function buildBatchFrameUrl(formPath) {
        const url = new window.URL(formPath, window.location.href);
        url.searchParams.set('office-export-mode', 'batch-frame');
        return url.toString();
    }

    function createBatchExportFrame(formPath) {
        const frame = document.createElement('iframe');
        frame.className = 'office-export-frame';
        frame.setAttribute('aria-hidden', 'true');
        frame.setAttribute('tabindex', '-1');
        frame.src = buildBatchFrameUrl(formPath);
        document.body.appendChild(frame);
        return frame;
    }

    function waitForFrameLoad(frame) {
        return new Promise(function (resolve, reject) {
            try {
                if (
                    frame.contentDocument &&
                    frame.contentDocument.readyState === 'complete' &&
                    frame.contentWindow &&
                    frame.contentWindow.location &&
                    frame.contentWindow.location.href !== 'about:blank'
                ) {
                    resolve(frame);
                    return;
                }
            } catch (error) {
                // Ignore early access errors and fall back to the load listener.
            }

            const timeoutId = window.setTimeout(function () {
                cleanup();
                reject(new Error('Timed out loading batch export frame.'));
            }, 30000);

            function cleanup() {
                window.clearTimeout(timeoutId);
                frame.removeEventListener('load', handleLoad);
                frame.removeEventListener('error', handleError);
            }

            function handleLoad() {
                cleanup();
                resolve(frame);
            }

            function handleError() {
                cleanup();
                reject(new Error('Failed loading batch export frame.'));
            }

            frame.addEventListener('load', handleLoad, { once: true });
            frame.addEventListener('error', handleError, { once: true });
        });
    }

    function requestFramePdfExport(frame, exportOptions) {
        return waitForFrameLoad(frame).then(function (loadedFrame) {
            return new Promise(function (resolve, reject) {
                const requestId = 'office-export-' + String(batchExportRequestCounter += 1);
                const timeoutId = window.setTimeout(function () {
                    cleanup();
                    reject(new Error('Timed out waiting for form export data.'));
                }, 180000);

                function cleanup() {
                    window.clearTimeout(timeoutId);
                    window.removeEventListener('message', handleMessage);
                }

                function handleMessage(event) {
                    const payload = event.data || {};

                    if (event.source !== loadedFrame.contentWindow || payload.requestId !== requestId) {
                        return;
                    }

                    if (payload.officeFormsType === 'office-forms-export-result') {
                        cleanup();
                        resolve({
                            filename: payload.filename,
                            data: payload.data
                        });
                        return;
                    }

                    if (payload.officeFormsType === 'office-forms-export-error') {
                        cleanup();
                        reject(new Error(payload.message || 'Form export failed.'));
                    }
                }

                window.addEventListener('message', handleMessage);
                loadedFrame.contentWindow.postMessage({
                    officeFormsType: 'office-forms-export-request',
                    requestId: requestId,
                    exportOptions: exportOptions
                }, '*');
            });
        });
    }

    function downloadBlob(blob, filename) {
        const objectUrl = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');

        anchor.href = objectUrl;
        anchor.download = filename;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();

        window.setTimeout(function () {
            window.URL.revokeObjectURL(objectUrl);
        }, 1000);
    }

    function exportManifestEntry(formPath, button, index, total) {
        const frame = createBatchExportFrame(formPath);
        button.textContent = 'Exporting ' + (index + 1) + '/' + total;

        return requestFramePdfExport(frame, {
            scale: BATCH_EXPORT_SCALE,
            compress: false
        })
            .then(function (result) {
                if (!result || !result.data || !result.filename) {
                    throw new Error('Form export returned no PDF data.');
                }

                return result;
            })
            .finally(function () {
                frame.remove();
            });
    }

    function exportAllForms(event) {
        event.preventDefault();

        const button = event.currentTarget;
        const labels = snapshotToolbarLabels();
        const manifest = getFormManifest();

        if (!manifest.length) {
            window.alert('No forms were found for Download All.');
            return;
        }

        disableToolbar();
        button.textContent = 'Preparing ZIP';

        ensureJsZip()
            .then(function (JSZip) {
                const zip = new JSZip();

                return manifest
                    .reduce(function (chain, formPath, index) {
                        return chain
                            .then(function () {
                                return exportManifestEntry(formPath, button, index, manifest.length);
                            })
                            .then(function (result) {
                                zip.file(result.filename, result.data);
                            });
                    }, Promise.resolve())
                    .then(function () {
                        button.textContent = 'Packing ZIP';
                        return zip.generateAsync({
                            type: 'blob',
                            compression: 'DEFLATE',
                            compressionOptions: { level: 6 }
                        });
                    });
            })
            .then(function (zipBlob) {
                downloadBlob(zipBlob, 'office-forms-all-pdfs.zip');
                restoreToolbar(labels);
            })
            .catch(function (error) {
                console.error('Download all failed', error);
                restoreToolbar(labels);
                window.alert('Download All failed. Please try again.');
            });
    }

    function exportPdf(event) {
        event.preventDefault();

        const button = event.currentTarget;
        const originalLabel = button.textContent;

        button.disabled = true;
        button.textContent = 'Preparing PDF';
        document.body.classList.add('office-forms-exporting');

        function restoreButton() {
            document.body.classList.remove('office-forms-exporting');
            button.disabled = false;
            button.textContent = originalLabel;
        }

        runPdfExport(null, function (worker) {
            return worker.save();
        })
            .then(function () {
                restoreButton();
            })
            .catch(function (error) {
                console.error('PDF export failed', error);
                restoreButton();
            });
    }

    function init() {
        if (!document.body || document.body.dataset.officeFormsEnhanced === 'true') {
            return;
        }

        document.body.dataset.officeFormsEnhanced = 'true';
        markFormIdentity();
        removeLegacyArtifacts();
        inlinePrimaryTitleBand();
        normalizeMasthead();
        markAvoidBreaks();
        ensureLogoDataUri();
        window.officeFormsExportPdfToArrayBuffer = exportCurrentPagePdfToArrayBuffer;
        window.officeFormsGetPdfFilename = getPdfFilename;
        setupBatchExportMessaging();

        if (isBatchFrameMode()) {
            return;
        }

        const buttons = normalizeExportButtons();
        buttons.download.addEventListener('click', exportPdf);
        buttons.batch.addEventListener('click', exportAllForms);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init, { once: true });
    } else {
        init();
    }
}());
