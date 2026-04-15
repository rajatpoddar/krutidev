(() => {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileInfo = document.getElementById('fileInfo');
  const fileName = document.getElementById('fileName');
  const fileSize = document.getElementById('fileSize');
  const removeFile = document.getElementById('removeFile');
  const convertBtn = document.getElementById('convertBtn');
  const progressSection = document.getElementById('progressSection');
  const progressFill = document.getElementById('progressFill');
  const progressPercent = document.getElementById('progressPercent');
  const successSection = document.getElementById('successSection');
  const downloadBtn = document.getElementById('downloadBtn');
  const convertAnother = document.getElementById('convertAnother');
  const errorMsg = document.getElementById('errorMsg');
  const errorText = document.getElementById('errorText');

  let selectedFile = null;
  let currentJobId = null;
  let pollInterval = null;

  // Format bytes
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function showError(msg) {
    errorText.textContent = msg;
    errorMsg.style.display = 'flex';
    setTimeout(() => { errorMsg.style.display = 'none'; }, 5000);
  }

  function hideError() {
    errorMsg.style.display = 'none';
  }

  function setFile(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
      showError('Invalid file type. Please upload an .xlsx or .xls file.');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      showError('File too large. Maximum size is 50MB.');
      return;
    }
    hideError();
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    fileInfo.style.display = 'flex';
    dropZone.style.display = 'none';
    convertBtn.disabled = false;
  }

  function resetUI() {
    selectedFile = null;
    currentJobId = null;
    if (pollInterval) clearInterval(pollInterval);
    fileInfo.style.display = 'none';
    dropZone.style.display = 'block';
    progressSection.style.display = 'none';
    successSection.style.display = 'none';
    convertBtn.disabled = true;
    convertBtn.style.display = 'flex';
    progressFill.style.width = '0%';
    progressPercent.textContent = '0%';
    fileInput.value = '';
    hideError();
  }

  // Drag & Drop
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  dropZone.addEventListener('click', (e) => {
    if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });

  removeFile.addEventListener('click', resetUI);

  convertAnother.addEventListener('click', resetUI);

  // Convert
  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    convertBtn.disabled = true;
    convertBtn.style.display = 'none';
    progressSection.style.display = 'block';
    hideError();

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      const res = await fetch('/upload', { method: 'POST', body: formData });
      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || 'Upload failed');
      }

      currentJobId = data.job_id;
      pollProgress(currentJobId);

    } catch (err) {
      progressSection.style.display = 'none';
      convertBtn.style.display = 'flex';
      convertBtn.disabled = false;
      showError(err.message || 'Something went wrong. Please try again.');
    }
  });

  function pollProgress(jobId) {
    pollInterval = setInterval(async () => {
      try {
        const res = await fetch(`/progress/${jobId}`);
        const data = await res.json();

        if (data.status === 'processing') {
          const pct = data.progress || 0;
          progressFill.style.width = pct + '%';
          progressPercent.textContent = pct + '%';

        } else if (data.status === 'done') {
          clearInterval(pollInterval);
          progressFill.style.width = '100%';
          progressPercent.textContent = '100%';

          setTimeout(() => {
            progressSection.style.display = 'none';
            successSection.style.display = 'block';
          }, 400);

        } else if (data.status === 'error') {
          clearInterval(pollInterval);
          progressSection.style.display = 'none';
          convertBtn.style.display = 'flex';
          convertBtn.disabled = false;
          showError(data.message || 'Conversion failed. Please try again.');
        }
      } catch {
        // Network hiccup, keep polling
      }
    }, 600);
  }

  // Download
  downloadBtn.addEventListener('click', () => {
    if (!currentJobId) return;
    window.location.href = `/download/${currentJobId}`;
  });
})();
