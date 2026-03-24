(function () {
    const CONTACT_HTML = [
        '<div class="doctors">',
        'Dr. Sambhav Jain <span>DMD, MS</span><br>',
        'Dr. Arpana Gupta <span>DDS, MDS</span>',
        '</div>',
        '<div class="contact-block"><strong>Email:</strong> info@bayareaimplantdentistry.com</div>',
        '<div class="contact-block"><strong>Main Office:</strong> 3381 Walnut Ave, Fremont, CA 94538<br><strong>Phone:</strong> (510) 574-0496</div>',
        '<div class="contact-block"><strong>SF Office:</strong> 4318 Geary Blvd, Suite 201, San Francisco, CA 94118<br><strong>Phone:</strong> (415) 696-2922</div>'
    ].join('');

    const PREMIUM_BAND_HTML = [
        '<div class="form-title">ALL<span class="bold-hyphen">-</span>ON<span class="bold-hyphen">-</span>8 ROBUST</div>',
        '<div class="form-label">Most Advanced Biomechanical Dental Implants Protocol</div>'
    ].join('');

    let logoDataUriPromise = null;

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

    function filenameFromTitle() {
        const cleanedTitle = (document.title || 'office-form')
            .replace(/\s*\|\s*All-On-8 Robust\s*$/i, '')
            .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '')
            .trim()
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-')
            .toLowerCase();

        return cleanedTitle || 'office-form';
    }

    function standardizeMasthead() {
        const masthead = document.querySelector('.masthead');
        if (!masthead) {
            return;
        }

        const mastheadRight = masthead.querySelector('.masthead-right');
        if (mastheadRight) {
            mastheadRight.innerHTML = CONTACT_HTML;
        }

        const logo = masthead.querySelector('.brand-logo img, img[src$="logo.svg"], img[src*="logo"]');
        if (logo) {
            logo.setAttribute('alt', 'Bay Area Implant Dentistry');
            logo.style.filter = 'grayscale(100%)';
        }
    }

    function standardizePremiumBand() {
        const masthead = document.querySelector('.masthead');
        let premiumBand = document.querySelector('.premium-product-band');

        if (!premiumBand && masthead) {
            premiumBand = document.createElement('div');
            premiumBand.className = 'premium-product-band';
            masthead.insertAdjacentElement('afterend', premiumBand);
        }

        if (premiumBand) {
            premiumBand.innerHTML = PREMIUM_BAND_HTML;
        }
    }

    function markAvoidBreaks() {
        document
            .querySelectorAll('.signature-section, .sig-section, .sig-row, .sig-block, .doc-footer')
            .forEach(function (node) {
                node.classList.add('office-print-avoid');
            });
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

    function addPageNumbers(pdf) {
        const pageCount = pdf.internal.getNumberOfPages();
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        pdf.setTextColor(0, 0, 0);
        pdf.setFont('times', 'normal');
        pdf.setFontSize(9);

        for (let page = 1; page <= pageCount; page += 1) {
            pdf.setPage(page);
            pdf.text('Page ' + page + ' of ' + pageCount, pageWidth - 0.35, pageHeight - 0.17, {
                align: 'right'
            });
        }
    }

    function buildPdfOptions(filename) {
        return {
            margin: [0.35, 0.35, 0.45, 0.35],
            filename: filename + '.pdf',
            image: { type: 'jpeg', quality: 1 },
            html2canvas: {
                scale: Math.min(3, Math.max(2.5, window.devicePixelRatio || 1)),
                useCORS: true,
                backgroundColor: '#ffffff',
                letterRendering: true,
                removeContainer: true,
                scrollX: 0,
                scrollY: 0
            },
            jsPDF: {
                unit: 'in',
                format: 'letter',
                orientation: 'portrait',
                compress: true
            },
            pagebreak: {
                mode: ['css', 'legacy'],
                avoid: ['.office-print-avoid', '.field-row', '.part-header', 'h3', '.checkbox-grid', '.conditions-container']
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

    function prepareExportResources(root) {
        const restorers = [];
        const pending = [];

        if (window.location.protocol === 'file:') {
            const inlineLogo = window.OFFICE_FORMS_LOGO_DATA_URI || '';

            root.querySelectorAll('img').forEach(function (image) {
                const parent = image.parentNode;
                if (!parent) {
                    return;
                }

                if (inlineLogo) {
                    const originalSrc = image.getAttribute('src');
                    image.setAttribute('src', inlineLogo);
                    pending.push(waitForImageReady(image));

                    restorers.push(function () {
                        image.setAttribute('src', originalSrc);
                    });
                    return;
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
                const worker = html2pdf()
                    .set(buildPdfOptions(filename))
                    .from(root)
                    .toPdf()
                    .get('pdf')
                    .then(function (pdf) {
                        addPageNumbers(pdf);
                    });

                return worker
                    .save()
                    .then(function () {
                        restoreButton(restoreResources);
                    })
                    .catch(function (error) {
                        console.error('PDF export failed', error);
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
        standardizeMasthead();
        standardizePremiumBand();
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
