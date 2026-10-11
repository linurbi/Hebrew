const page = window.PAGE;
const fileInput = document.querySelector('#file');
const chosen = document.querySelector('#chosen');
const button = document.querySelector('#go');
const status = document.querySelector('#status');
const drop = document.querySelector('#drop');
const contactStatus = document.querySelector('#contact-status');

function track(name) {
  if (window.HebrewStats) HebrewStats.event(name);
}

function showError(text) {
  status.className = 'status error';
  status.textContent = text;
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  chosen.textContent = file ? file.name : page.chooseHint;
  button.disabled = !file;
  status.replaceChildren();
});

['dragover', 'dragenter'].forEach((name) => drop.addEventListener(name, (event) => {
  event.preventDefault();
  drop.classList.add('over');
}));
['dragleave', 'drop'].forEach((name) => drop.addEventListener(name, (event) => {
  event.preventDefault();
  drop.classList.remove('over');
}));
drop.addEventListener('drop', (event) => {
  if (!event.dataTransfer.files.length) return;
  fileInput.files = event.dataTransfer.files;
  fileInput.dispatchEvent(new Event('change'));
});

document.querySelector('#form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) {
    track('fix-rejected');
    showError(page.errors.rejected);
    return;
  }
  if (file.size > 20 * 1024 * 1024) {
    track('fix-rejected');
    showError(page.errors['too-big']);
    return;
  }
  button.disabled = true;
  status.className = 'status';
  status.textContent = page.fixing;
  try {
    const result = await HebrewFix.fixDocx(await file.arrayBuffer());
    status.replaceChildren();
    if (!result.ok) {
      track('fix-failed');
      showError(page.errors[result.error] || page.errors.failed);
      button.disabled = false;
      return;
    }
    track(result.warnings.length ? 'fix-needs-review' : 'fix-success');
    const url = URL.createObjectURL(new Blob([result.bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }));
    status.className = 'status';
    status.textContent = page.ready;
    const link = document.createElement('a');
    link.className = 'download';
    link.href = url;
    link.download = file.name.replace(/\.docx$/i, page.suffix);
    link.textContent = page.download;
    status.append(link);
    if (result.warnings.length) {
      const box = document.createElement('div');
      box.className = 'warnings';
      const title = document.createElement('strong');
      title.textContent = page.review;
      box.append(title);
      result.warnings.forEach((warning) => {
        const line = document.createElement('p');
        const reasons = warning.reasons.map((reason) => page.reasons[reason] || reason);
        line.textContent = warning.text + ' ' + reasons.join(' ');
        box.append(line);
      });
      status.append(box);
    }
  } catch (error) {
    track('fix-failed');
    showError(page.errors.failed);
  }
  button.disabled = false;
});

document.querySelector('#contact').addEventListener('submit', async (event) => {
  event.preventDefault();
  const form = event.currentTarget;
  const send = form.querySelector('button');
  if (form.elements.company.value) return;
  send.disabled = true;
  contactStatus.className = 'status';
  contactStatus.textContent = page.sending;
  try {
    const response = await fetch('https://formsubmit.co/ajax/linurbi@gmail.com', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify({
        name: form.elements.name.value,
        email: form.elements.email.value,
        message: form.elements.message.value,
        _subject: 'DocFixes contact',
        _captcha: 'false',
        _template: 'table',
      }),
    });
    if (!response.ok) throw new Error('contact failed');
    form.reset();
    contactStatus.className = 'status';
    contactStatus.textContent = page.sent;
  } catch (error) {
    contactStatus.className = 'status error';
    contactStatus.textContent = page.sendFailed;
  }
  send.disabled = false;
});
