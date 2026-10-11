/* Privacy-friendly counts. Set ENDPOINT to the GoatCounter count URL.
   Only event names are sent. File names and document contents are never sent. */
(function () {
  const ENDPOINT = '';
  const queue = [];

  function emit(name) {
    if (!ENDPOINT || !name) return;
    if (window.goatcounter && window.goatcounter.count) {
      window.goatcounter.count({ path: name, event: true, no_session: true });
      return;
    }
    queue.push(name);
  }

  function flush() {
    if (!window.goatcounter || !window.goatcounter.count) return;
    while (queue.length) emit(queue.shift());
  }

  if (ENDPOINT) {
    const script = document.createElement('script');
    script.async = true;
    script.src = 'https://gc.zgo.at/count.js';
    script.dataset.goatcounter = ENDPOINT;
    script.onload = flush;
    document.head.appendChild(script);
  }

  window.HebrewStats = { event: emit };
})();
