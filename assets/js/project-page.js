/* Shared page bootstrap. Only approved configuration entries become UI links. */
(() => {
  const project = globalThis.PROJECTS_CONFIG?.projects.find(item => item.id === document.body.dataset.project);
  if (!project || !globalThis.ExcelScreenshotViewer) return;
  document.querySelectorAll('[data-excel-viewer]').forEach(host => {
    const group = project[host.dataset.excelViewer];
    if (group?.type === 'excel') {
      // Restricted previews can still have approved PNGs, but no workbook download.
      const preview = group.workbook?.downloadAvailable ? group : {
        ...group,
        sheets: group.sheets.map(sheet => ({ ...sheet, coverageNote: 'Read-only screenshot preview. Some output previews are currently unavailable.' }))
      };
      new globalThis.ExcelScreenshotViewer(host, preview, project.title);
    }
  });
  document.querySelectorAll('[data-pdf-viewer]').forEach(host => {
    const group = project[host.dataset.pdfViewer];
    if (group?.type === 'pdf' && globalThis.NativePdfViewer) new globalThis.NativePdfViewer(host, group, project.title);
  });
  document.querySelectorAll('[data-project-downloads]').forEach(host => {
    ['raw', 'output'].forEach(kind => {
      const workbook = project[kind]?.workbook;
      if (!workbook?.downloadAvailable || !workbook.path?.startsWith('/')) return;
      const link = document.createElement('a');
      link.className = 'button';
      link.href = workbook.path;
      link.download = workbook.filename;
      link.textContent = `${kind === 'raw' ? 'Input' : 'Report'}: ${workbook.filename}`;
      host.append(link);
    });
  });
})();
