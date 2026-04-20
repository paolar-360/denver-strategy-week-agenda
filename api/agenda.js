module.exports = async (req, res) => {
  try {
    // Lazy-require so we can surface module-load errors in the response body
    // instead of crashing the whole serverless invocation at import time.
    const { generateBuffer } = require('../agenda_branded.js');
    const buf = await generateBuffer();
    res.statusCode = 200;
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="Strategy_Week_Agenda.docx"'
    );
    res.setHeader('Content-Length', buf.length);
    res.setHeader('Cache-Control', 'public, max-age=300, s-maxage=300');
    res.end(buf);
  } catch (err) {
    console.error('agenda error:', err);
    res.statusCode = 500;
    res.setHeader('Content-Type', 'application/json');
    res.end(
      JSON.stringify({
        error: 'Failed to generate document',
        message: err && err.message ? err.message : String(err),
        stack: err && err.stack ? String(err.stack).split('\n').slice(0, 6) : undefined
      })
    );
  }
};
