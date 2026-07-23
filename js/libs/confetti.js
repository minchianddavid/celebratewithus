const zIndex = 1057;
let openingCanvas = null;
let openingFire = null;

/**
 * @returns {any}
 */
const heartShape = () => {
    return window.confetti.shapeFromPath({
        path: 'M167 72c19,-38 37,-56 75,-56 42,0 76,33 76,75 0,76 -76,151 -151,227 -76,-76 -151,-151 -151,-227 0,-42 33,-75 75,-75 38,0 57,18 76,56z',
        matrix: [0.03333333333333333, 0, 0, 0.03333333333333333, -5.566666666666666, -5.533333333333333]
    });
};

/**
 * @returns {void}
 */
export const prepareBasicAnimation = () => {
    if (!window.confetti || openingFire) {
        return;
    }

    openingCanvas = document.createElement('canvas');
    openingCanvas.className = 'opening-confetti-canvas';
    openingCanvas.setAttribute('aria-hidden', 'true');
    openingCanvas.width = window.innerWidth;
    openingCanvas.height = window.innerHeight;
    document.body.appendChild(openingCanvas);

    openingFire = window.confetti.create(openingCanvas, {
        resize: false,
        useWorker: true,
    });

    // Warm the worker and canvas while the loading page is still visible.
    openingFire({
        particleCount: 1,
        startVelocity: 0,
        ticks: 1,
        gravity: 0,
        scalar: 0,
        origin: { y: 1 },
        disableForReducedMotion: true,
    }).catch(() => {});
};

/**
 * @returns {Promise<void>}
 */
export const basicAnimation = () => {
    prepareBasicAnimation();
    if (!openingFire || !openingCanvas) {
        return Promise.resolve();
    }

    const fire = openingFire;
    const canvas = openingCanvas;
    openingFire = null;
    openingCanvas = null;

    return fire({
        particleCount: 544,
        startVelocity: 54,
        spread: 172,
        ticks: 384,
        gravity: 0.78,
        scalar: 0.56,
        origin: { y: 1 },
        disableForReducedMotion: true,
    }).finally(() => canvas.remove());
};

/**
 * @param {number} [until=15]
 * @returns {void}
 */
export const openAnimation = (until = 15) => {
    if (!window.confetti) {
        return;
    }

    const duration = until * 1000;
    const animationEnd = Date.now() + duration;

    const heart = heartShape();
    const colors = ['#FFC0CB', '#C71585'];

    const randomInRange = (min, max) => {
        return Math.random() * (max - min) + min;
    };

    const frame = () => {
        const timeLeft = animationEnd - Date.now();

        colors.forEach((color) => {
            const fromLeft = Math.random() < 0.5;
            window.confetti({
                particleCount: 1,
                startVelocity: 0,
                ticks: Math.max(50, 75 * (timeLeft / duration)),
                origin: {
                    x: fromLeft ? randomInRange(0, 0.22) : randomInRange(0.78, 1),
                    y: Math.abs(Math.random() - (timeLeft / duration)),
                },
                zIndex: zIndex,
                colors: [color],
                shapes: [heart],
                drift: randomInRange(-0.5, 0.5),
                gravity: randomInRange(0.35, 0.7),
                scalar: randomInRange(0.45, 0.8),
            });
        });

        if (timeLeft > 0) {
            requestAnimationFrame(frame);
        }
    };

    requestAnimationFrame(frame);
};

/**
 * @param {HTMLElement} div
 * @param {number} [duration=50]
 * @returns {void}
 */
export const tapTapAnimation = (div, duration = 50) => {
    if (!window.confetti) {
        return;
    }

    const end = Date.now() + duration;
    const domRec = div.getBoundingClientRect();
    const yPosition = Math.max(0.3, Math.min(1, (domRec.top / window.innerHeight) + 0.2));

    const heart = heartShape();
    const colors = ['#FF69B4', '#FF1493'];

    const frame = () => {
        colors.forEach((color) => {
            window.confetti({
                particleCount: 2,
                angle: 60,
                spread: 55,
                shapes: [heart],
                origin: { x: domRec.left / window.innerWidth, y: yPosition },
                zIndex: zIndex,
                colors: [color]
            });
            window.confetti({
                particleCount: 2,
                angle: 120,
                spread: 55,
                shapes: [heart],
                origin: { x: domRec.right / window.innerWidth, y: yPosition },
                zIndex: zIndex,
                colors: [color]
            });
        });

        if (Date.now() < end) {
            requestAnimationFrame(frame);
        }
    };

    requestAnimationFrame(frame);
};

/**
 * @param {HTMLElement} element
 * @param {'left'|'right'} sourceSide
 * @returns {void}
 */
export const journeyGoldAnimation = (element, sourceSide = 'right') => {
    if (!window.confetti) {
        return;
    }

    const rect = element.getBoundingClientRect();
    const fromRight = sourceSide === 'right';
    const centerX = (rect.left + (rect.width / 2)) / window.innerWidth;
    const titleY = (rect.top + (rect.height * 0.42)) / window.innerHeight;
    window.confetti({
        particleCount: 38,
        angle: fromRight ? 155 : 25,
        spread: 92,
        startVelocity: 26,
        ticks: 78,
        gravity: 0.92,
        scalar: 0.58,
        shapes: ['star'],
        origin: {
            x: Math.max(0.05, Math.min(0.95, (fromRight ? rect.right - 6 : rect.left + 6) / window.innerWidth)),
            y: Math.max(0.1, Math.min(0.88, titleY)),
        },
        zIndex: zIndex,
        colors: ['#b8976a', '#d6bd91', '#eadfc9', '#f8f0dc', '#9f7b4d'],
        disableForReducedMotion: true,
    });

    window.setTimeout(() => {
        window.confetti({
            particleCount: 16,
            angle: 270,
            spread: 116,
            startVelocity: 11,
            ticks: 62,
            gravity: 1.08,
            scalar: 0.46,
            shapes: ['star'],
            origin: {
                x: Math.max(0.08, Math.min(0.92, centerX)),
                y: Math.max(0.1, Math.min(0.88, titleY)),
            },
            zIndex: zIndex,
            colors: ['#d6bd91', '#eadfc9', '#fff8e8'],
            disableForReducedMotion: true,
        });
    }, 140);
};

/**
 * A restrained shower for the RSVP confirmation card.
 * @returns {void}
 */
export const rsvpPetalAnimation = () => {
    if (!window.confetti) {
        return;
    }

    const petal = window.confetti.shapeFromPath({
        path: 'M10 0 C16 4 18 12 10 20 C2 12 4 4 10 0 Z',
        matrix: [0.5, 0, 0, 0.5, -5, -5],
    });

    window.confetti({
        particleCount: 7,
        angle: 270,
        spread: 42,
        startVelocity: 5,
        ticks: 48,
        gravity: 0.38,
        drift: 0.08,
        scalar: 0.52,
        origin: { x: 0.5, y: 0.04 },
        zIndex: zIndex,
        shapes: [petal],
        colors: ['#b8976a', '#d6bd91', '#eadfc9'],
        disableForReducedMotion: true,
    });
};

/**
 * @param {HTMLElement} element
 * @param {boolean} subtle
 * @returns {void}
 */
export const loveSparkleAnimation = (element, subtle = false) => {
    if (!window.confetti) {
        return;
    }

    const rect = element.getBoundingClientRect();
    const origin = {
        x: (rect.left + (rect.width / 2)) / window.innerWidth,
        y: (rect.top + (rect.height / 2)) / window.innerHeight,
    };
    const burst = {
        particleCount: subtle ? 3 : 7,
        spread: subtle ? 24 : 34,
        startVelocity: subtle ? 10 : 17,
        ticks: subtle ? 42 : 58,
        gravity: subtle ? 0.48 : 0.62,
        scalar: subtle ? 0.44 : 0.6,
        shapes: ['star'],
        origin: origin,
        zIndex: zIndex,
        colors: ['#b8976a', '#d6bd91', '#f1e8d5', '#9f7b4d'],
    };

    window.confetti({ ...burst, angle: 155 });
    window.confetti({ ...burst, angle: 25 });
};

/**
 * A wide burst of fine champagne-gold dust for the Hero photo spin.
 * @param {HTMLElement} element
 * @returns {void}
 */
export const heroSpinGoldDustAnimation = (element) => {
    if (!window.confetti) {
        return;
    }

    const rect = element.getBoundingClientRect();
    window.confetti({
        particleCount: 46,
        spread: 112,
        startVelocity: 19,
        ticks: 92,
        gravity: 0.4,
        scalar: 0.34,
        shapes: ['circle'],
        origin: {
            x: (rect.left + (rect.width / 2)) / window.innerWidth,
            y: (rect.top + (rect.height / 2)) / window.innerHeight,
        },
        zIndex: zIndex,
        colors: ['#b99052', '#d7ad62', '#f0d38f', '#fff8df', '#f8edd0'],
    });
};
