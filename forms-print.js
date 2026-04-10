(function () {
    let logoDataUriPromise = null;
    const monochromeLogoCache = new Map();

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

    function normalizeDownloadButton() {
        let button = document.getElementById('downloadPdfBtn');

        if (!button) {
            button = document.createElement('button');
            button.id = 'downloadPdfBtn';
            button.className = 'fab';
            button.type = 'button';
            document.body.appendChild(button);
        }

        const cleanButton = button.cloneNode(false);
        cleanButton.id = 'downloadPdfBtn';
        cleanButton.className = 'fab';
        cleanButton.type = 'button';
        cleanButton.textContent = 'Download PDF';

        button.replaceWith(cleanButton);
        return cleanButton;
    }

    function buildPdfOptions(filename) {
        return {
            margin: [0.18, 0.18, 0.18, 0.18],
            filename: filename + '.pdf',
            image: { type: 'jpeg', quality: 1 },
            html2canvas: {
                scale: Math.min(3.4, Math.max(2.8, window.devicePixelRatio || 1)),
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
                    '.delivery-note'
                ]
            }
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

        if (window.location.protocol === 'file:') {
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
        }

        return Promise.all(pending).then(function () {
            return function restoreResources() {
                while (restorers.length) {
                    restorers.pop()();
                }
            };
        });
    }

    function exportPdf(event) {
        event.preventDefault();

        const button = event.currentTarget;
        const root = document.getElementById('pdf-content') || document.querySelector('.document');

        if (!root || typeof window.html2pdf !== 'function') {
            return;
        }

        const originalLabel = button.textContent;
        const filename = filenameFromTitle();
        const pdfOptions = buildPdfOptions(filename);

        button.disabled = true;
        button.textContent = 'Preparing PDF';
        document.body.classList.add('office-forms-exporting');
        window.scrollTo(0, 0);

        function restoreButton(restoreResources) {
            if (typeof restoreResources === 'function') {
                restoreResources();
            }
            document.body.classList.remove('office-forms-exporting');
            button.disabled = false;
            button.textContent = originalLabel;
        }

        ensureLogoDataUri()
            .then(function () {
                return prepareExportResources(root);
            })
            .then(function (restoreResources) {
                const restorePagination = preparePagination(root, pdfOptions);
                const worker = html2pdf()
                    .set(pdfOptions)
                    .from(root)
                    .save();

                return worker
                    .then(function () {
                        restorePagination();
                        restoreButton(restoreResources);
                    })
                    .catch(function (error) {
                        console.error('PDF export failed', error);
                        restorePagination();
                        restoreButton(restoreResources);
                    });
            })
            .catch(function (error) {
                console.error('PDF export setup failed', error);
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

        const button = normalizeDownloadButton();
        button.addEventListener('click', exportPdf);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init, { once: true });
    } else {
        init();
    }
}());
