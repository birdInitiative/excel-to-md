const express = require('express');
const multer = require('multer');
const showdown = require('showdown');
const converter = new showdown.Converter();
const xlsx = require('xlsx');
const axios = require('axios');
const markdownToAst = require('markdown-to-ast');

const app = express();
const port = 3000;

const storage = multer.memoryStorage();  // Store the file in memory
const upload = multer({ storage: storage });

const OPENAI_API_ENDPOINT = 'https://api.openai.com/v1/engines/davinci/completions';
const OPENAI_API_KEY = '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';

app.use(express.static('public'));  // Serve static files

async function generateIndexFromMarkdown(markdown) {
    const headers = {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json',
    };

    // Extract headings from markdown
    const ast = markdownToAst.parse(markdown);
    const headings = ast.children.filter(node => node.type === 'Header').map(node => node.children[0].value);

    const data = {
        prompt: `Generate an index like book from the following text:\n\n${headings.join('\n')}`,
        max_tokens: 1500000
    };

    try {
        const response = await axios.post(OPENAI_API_ENDPOINT, data, { headers: headers });
        return response.data.choices[0].text.trim();
    } catch (error) {
        throw error;
    }
}

app.post('/convert', upload.single('excelFile'), async (req, res) => {
    try {
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });  // Get data as an array of arrays
        let markdownOutput = '';

        jsonData.forEach(row => {
            row.forEach((cell, index) => {
                if (cell && typeof cell === 'string' && cell.trim() !== '') {
                    const indentation = '  '.repeat(index); // Determine indentation based on column index
                    const isParent = row[index + 1] && typeof row[index + 1] === 'string' && row[index + 1].trim() !== '';
                    if (isParent) {
                        markdownOutput += `${indentation}- ${cell.trim()}\n`;
                    } else {
                        markdownOutput += `${indentation}  - ${cell.trim()}\n`;
                    }
                }
            });
            markdownOutput += '\n'; // Add a newline between rows for clarity
        });

        // Render markdown as HTML using Showdown
        const renderedMarkdown = converter.makeHtml(markdownOutput);

        res.send(renderedMarkdown);

    } catch (error) {
        console.error('Error Details:', error.message, error.stack);  
        res.status(500).send('Error processing Excel file. Please check server logs for details.');
    }
});

app.listen(port, () => {
    console.log(`Server started on http://localhost:${port}`);
});

