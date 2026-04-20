const { generateBuffer } = require('../agenda_branded.js');

module.exports = async (req, res) => {
  try {
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
    console.error(err);
    res.statusCode = 500;
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify({ error: 'Failed to generate document' }));
  }
};
