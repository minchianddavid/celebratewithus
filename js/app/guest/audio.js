import { progress } from './progress.js';
import { util } from '../../common/util.js';
import { cache } from '../../connection/cache.js';

export const audio = (() => {

    const statePlay = '<i class="fa-solid fa-circle-pause"></i>';
    const statePause = '<i class="fa-solid fa-circle-play"></i>';

    /**
     * @param {boolean} [playOnOpen=true]
     * @returns {Promise<void>}
     */
    const load = async (playOnOpen = true) => {

        const url = document.body.getAttribute('data-audio');
        if (!url) {
            progress.complete('audio', true);
            return;
        }

        /**
         * @type {HTMLAudioElement|null}
         */
        let audioEl = null;

        try {
            const audioUrl = await cache('audio').withForceCache().get(url, progress.getAbort());
            audioEl = new Audio();
            audioEl.loop = true;
            audioEl.muted = false;
            audioEl.autoplay = false;
            audioEl.controls = false;
            audioEl.preload = 'auto';

            await new Promise((resolve, reject) => {
                let readyTimer = null;
                let cleanup = () => {};
                const ready = () => {
                    cleanup();
                    resolve();
                };
                const fail = () => {
                    cleanup();
                    reject(new Error('Audio preload failed'));
                };
                cleanup = () => {
                    window.clearTimeout(readyTimer);
                    audioEl.removeEventListener('canplaythrough', ready);
                    audioEl.removeEventListener('error', fail);
                };

                audioEl.addEventListener('canplaythrough', ready, { once: true });
                audioEl.addEventListener('error', fail, { once: true });
                readyTimer = window.setTimeout(ready, 5000);
                audioEl.src = audioUrl;
                audioEl.load();

                if (audioEl.readyState >= 4) {
                    ready();
                }
            });

            progress.complete('audio');
        } catch {
            progress.invalid('audio');
            return;
        }

        let isPlay = false;
        const music = document.getElementById('button-music');

        /**
         * @returns {Promise<void>}
         */
        const play = async () => {
            if (!navigator.onLine || !music) {
                return;
            }

            music.disabled = true;
            try {
                await audioEl.play();
                isPlay = true;
                music.disabled = false;
                music.innerHTML = statePlay;
            } catch (err) {
                isPlay = false;
                util.notify(err).error();
            }
        };

        /**
         * @returns {void}
         */
        const pause = () => {
            isPlay = false;
            audioEl.pause();
            music.innerHTML = statePause;
        };

        document.addEventListener('undangan.open', () => {
            music.classList.remove('d-none');

            if (playOnOpen) {
                play();
            }
        });

        music.addEventListener('offline', pause);
        music.addEventListener('click', () => isPlay ? pause() : play());

        // Auto-pause when user leaves the page (tab switch, minimize, lock phone)
        let wasPlaying = false;
        document.addEventListener('visibilitychange', () => {
            if (document.hidden && isPlay) {
                wasPlaying = true;
                pause();
            } else if (!document.hidden && wasPlaying) {
                wasPlaying = false;
                play();
            }
        });
    };

    /**
     * @returns {object}
     */
    const init = () => {
        progress.add();

        return {
            load,
        };
    };

    return {
        init,
    };
})();
