const fs = require('fs');
const path = require('path');

class CustomHtmlReporter {
  constructor() {
    this.results = [];
    this.reportDir = path.resolve(__dirname, '../report');
    if (!fs.existsSync(this.reportDir)) fs.mkdirSync(this.reportDir, { recursive: true });
  }

  onBegin(config) {
    this.browserName = config.projects[0].use.browserName;
  }

  onTestEnd(test, result) {
    const scenarioId = test.title.replace(/\s+/g, '_').toLowerCase();
    const duration = result.duration;
    const statusEmoji = result.status === 'passed' ? '✅️' : result.status === 'failed' ? '⛔' : '⚠️';
    const browserEmoji = this.browserName === 'chromium' ? '🌐' : this.browserName === 'webkit' ? '🧭' : '🦊';

    const passedSteps = result.steps.filter(s => s.error === undefined && s.category === 'test.step').length;
    const failedSteps = result.steps.filter(s => s.error !== undefined && s.category === 'test.step').length;
    const skippedSteps = result.steps.filter(s => s.category === 'test.step' && s.status === 'skipped').length;
    const totalSteps = passedSteps + failedSteps + skippedSteps;

    const scenario = {
      id: scenarioId,
      name: test.title,
      status: result.status,
      statusEmoji,
      browserEmoji,
      totalSteps,
      passedSteps,
      failedSteps,
      skippedSteps,
      duration,
      steps: result.steps.map(s => ({
        title: s.title,
        error: s.error,
        duration: s.duration,
      })),
      error: result.error?.message,
      screenshot: result.attachments.find(a => a.name === 'screenshot')?.path,
      video: result.attachments.find(a => a.name === 'video')?.path
    };

    this.results.push(scenario);

    // Generate individual detail page
    this.generateScenarioPage(scenario);
  }

  onEnd() {
    this.generateSummaryPage();
  }

generateScenarioPage(scenario) {
  const scenarioFile = path.join(this.reportDir, `${scenario.id}.html`);

  // Copy screenshot to report folder
  let screenshotTag = '';
  if (scenario.screenshot && fs.existsSync(scenario.screenshot)) {
    const destScreenshot = path.join(this.reportDir, `${scenario.id}_screenshot.png`);
    fs.copyFileSync(scenario.screenshot, destScreenshot);
    screenshotTag = `<h3>📸 Screenshot</h3><img src="${path.basename(destScreenshot)}" width="600" />`;
  }

  // Copy video to report folder
  let videoTag = '';
  if (scenario.video && fs.existsSync(scenario.video)) {
    const destVideo = path.join(this.reportDir, `${scenario.id}_video.webm`);
    fs.copyFileSync(scenario.video, destVideo);
    videoTag = `<h3>🎥 Video</h3><video src="${path.basename(destVideo)}" width="600" controls></video>`;
  }

  const filteredSteps = scenario.steps.filter(step =>
    ![
      'Before Hooks',
      'After Hooks',
      'Worker Cleanup',
      "locator.clear(//input[@id='last-name'])",
      "locator.clear(//input[@id='postal-code'])"
    ].includes(step.title)
  );

  const html = `
    <html>
      <head>
        <title>${scenario.name}</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
          }
          h2, h3 {
            color: #444;
          }
          ul {
            padding-left: 20px;
          }
          li {
            margin-bottom: 8px;
          }
          a {
            text-decoration: none;
            color: #007bff;
          }
          a:hover {
            text-decoration: underline;
          }
        </style>
      </head>
      <body>
        <h2>${scenario.name}</h2>
        <a href="index.html">⬅️ Back to Summary</a>
        ${scenario.status === 'failed' ? `<h3 style="color:red;">❌ Reason: ${scenario.error}</h3>` : ''}
        <h3>📝 Steps:</h3>
        <ul>
          ${filteredSteps.map(step => `
            <li>${step.title} ${step.error ? `<span style="color:red;">(Failed)</span>` : ''}</li>
          `).join('')}
        </ul>
        ${screenshotTag}
        ${videoTag}
      </body>
    </html>
  `;

  fs.writeFileSync(scenarioFile, html, 'utf-8');
}


generateSummaryPage() {
  const total = this.results.length;
  const passed = this.results.filter(r => r.status === 'passed').length;
  const failed = this.results.filter(r => r.status === 'failed').length;
  const skipped = this.results.filter(r => r.status === 'skipped').length;

  const chartData = `<script>
    const ctx = document.getElementById('chart').getContext('2d');
    new Chart(ctx, {
      type: 'pie',
      data: {
        labels: ['Passed', 'Failed', 'Skipped'],
        datasets: [{
          data: [${passed}, ${failed}, ${skipped}],
          backgroundColor: ['#4CAF50', '#F44336', '#FF9800']
        }]
      },
      options: { responsive: false }
    });
  </script>`;

  const tableRows = this.results.map(r => {
    return `
      <tr>
        <td><a href="${r.id}.html">${r.name}</a></td>
        <td>${r.statusEmoji}</td>
        <td>1</td>
        <td>${r.status === 'passed' ? 1 : 0}</td>
        <td>${r.status === 'failed' ? 1 : 0}</td>
        <td>${r.status === 'skipped' ? 1 : 0}</td>
        <td>${r.duration}</td>
      </tr>
    `;
  }).join('');

  const html = `
    <html>
      <head>
        <title>Playwright Test Execution Report</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
          }
          h1, h2 {
            color: #333;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
          }
          th, td {
            padding: 10px;
            text-align: center;
            border-bottom: 1px solid #ddd;
          }
          th {
            background-color: #f4f4f4;
          }
          tr:hover {
            background-color: #f1f1f1;
          }
          a {
            text-decoration: none;
            color: #007bff;
            font-weight: bold;
          }
          a:hover {
            text-decoration: underline;
          }
        </style>
      </head>
      <body>
        <h1>📊 Playwright Test Execution Report</h1>
        <h2>Summary</h2>
        <p><strong>Total Scenarios:</strong> ${total} | ✅ Passed: ${passed} | ❌ Failed: ${failed} | ⚠️ Skipped: ${skipped}</p>
        <canvas id="chart" width="200" height="200"></canvas>
        ${chartData}
        <h2>Detailed Results</h2>
        <table>
          <thead>
            <tr>
              <th>Test Scenario Name</th>
              <th>Status</th>
              <th>Total</th>
              <th>Passed</th>
              <th>Failed</th>
              <th>Skipped</th>
              <th>Duration (ms)</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>
      </body>
    </html>
  `;

  fs.writeFileSync(path.join(this.reportDir, 'index.html'), html, 'utf-8');
}


}

module.exports = CustomHtmlReporter;
