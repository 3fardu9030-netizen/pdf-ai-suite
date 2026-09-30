const express = require('express');
const router = express.Router();
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const jwt = require('jsonwebtoken');
const { PDFParse } = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const mammoth = require('mammoth');
const PDFDocument = require('pdfkit');
const { PDFDocument: PDFLibDocument } = require('pdf-lib');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const DocumentModel = require('../models/Document');

const JWT_SECRET = 'your_super_secret_jwt_key_here';

// Auth Middleware
function verifyToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];
  
  if (!token) {
    return res.status(401).json({ success: false, error: 'Access denied. No token provided.' });
  }

  try {
    const verified = jwt.verify(token, JWT_SECRET);
    req.user = verified;
    next();
  } catch (err) {
    res.status(403).json({ success: false, error: 'Invalid token.' });
  }
}

// Configure multer storage with support for documents and images
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage: storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = /pdf|docx|pptx|png|jpg|jpeg/;
    const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = allowedTypes.test(file.mimetype) || file.mimetype.startsWith('image/');
    if (extname || mimetype) {
      return cb(null, true);
    }
    cb(new Error('Only PDF, Word, PowerPoint, and Image files are allowed!'));
  }
});

async function saveHistoryRecord(userId, title, fileUrl, fileType) {
  try {
    await DocumentModel.create({ userId, title, fileType, fileUrl });
  } catch (err) {
    console.error('Failed to save document history record:', err);
  }
}
function parseTargetBytes(value) {
  const match = String(value || '').trim().match(
    /^(\d+(?:\.\d+)?)\s*(B|KB|MB)$/i
  );

  if (!match) return null;

  const amount = Number(match[1]);
  const unit = match[2].toUpperCase();
  const multiplier = {
    B: 1,
    KB: 1024,
    MB: 1024 * 1024
  };

  if (!Number.isFinite(amount) || amount <= 0) return null;

  return Math.floor(amount * multiplier[unit]);
}

function sizeReport(originalBytes, newBytes, targetBytes) {
  return {
    originalSize: originalBytes,
    newSize: newBytes,
    targetSize: targetBytes,
    percentageChange: Number(
      (((newBytes - originalBytes) / originalBytes) * 100).toFixed(2)
    )
  };
}

// Helper: Automatically chunk bullets across multiple slides to prevent overcrowding
function addChunksToSlides(pptx, title, bullets) {
  const MAX_BULLETS_PER_SLIDE = 5;
  if (!bullets || bullets.length === 0) {
    let slide = pptx.addSlide();
    slide.addText(title, { x: 0.8, y: 0.5, w: 11.5, h: 0.8, fontSize: 22, bold: true, color: '003366' });
    return;
  }

  for (let i = 0; i < bullets.length; i += MAX_BULLETS_PER_SLIDE) {
    const chunk = bullets.slice(i, i + MAX_BULLETS_PER_SLIDE);
    let slide = pptx.addSlide();
    
    let slideTitle = i === 0 ? title : `${title} (Cont.)`;
    
    slide.addText(slideTitle, {
      x: 0.8,
      y: 0.5,
      w: 11.5,
      h: 0.8,
      fontSize: 22,
      bold: true,
      color: '003366'
    });

    const bulletObjects = chunk.map(b => ({
      text: b,
      options: { fontSize: 13, color: '333333', bullet: true, spaceAfter: 8 }
    }));

    slide.addText(bulletObjects, {
      x: 0.8,
      y: 1.4,
      w: 11.5,
      h: 5.2,
      lineSpacing: 16
    });
  }
}

// 1. AI PDF Generator
router.post('/generate-ai', verifyToken, async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ success: false, error: 'Prompt is required' });
    
    const mockResultText = `Generated content based on prompt: "${prompt}"`;
    const outputFilename = 'ai-generated-' + Date.now() + '.pdf';
    const outputPath = path.join('uploads', outputFilename);

    const doc = new PDFDocument();
    const stream = fs.createWriteStream(outputPath);
    doc.pipe(stream);
    doc.fontSize(20).text('AI Generated Document', { align: 'center' });
    doc.moveDown();
    doc.fontSize(12).text(mockResultText);
    doc.end();

    stream.on('finish', async () => {
      const fileUrl = `/uploads/${outputFilename}`;
      const title = 'AI: ' + prompt.substring(0, 30) + '...';
      await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
      res.json({ success: true, result: mockResultText, fileUrl, title });
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// PDF Summarizer
router.post('/summarize-pdf', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    
    const dataBuffer = fs.readFileSync(req.file.path);
    const parser = new PDFParse({ data: dataBuffer });
    const pdfData = await parser.getText();
    await parser.destroy();
    
    const extractedText = pdfData.text || '';
    if (!extractedText.trim()) {
      return res.status(400).json({ success: false, error: 'No extractable text found in this PDF.' });
    }

    const cleanText = extractedText.replace(/\s+/g, ' ').trim();
    const words = cleanText.split(' ');
    const totalWords = words.length;
    
    const previewLength = Math.min(150, words.length);
    const sampleSnippet = words.slice(0, previewLength).join(' ');

    const summaryText = `[PDF Summary Overview]\n\n• Document Name: ${req.file.originalname}\n• Total Length: ~${totalWords} words\n\nKey Content Insights:\n"${sampleSnippet}..."\n\nSummary Statement: This document contains approximately ${totalWords} words focusing on core thematic elements derived from its introduction and text sections.`;

    const title = `${req.file.originalname} (Summary)`;
    await saveHistoryRecord(req.user.userId, title, '#', 'summary');

    res.json({ success: true, summary: summaryText });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 2. PDF to Word
router.post('/pdf-to-word', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    const dataBuffer = fs.readFileSync(req.file.path);
    const parser = new PDFParse({ data: dataBuffer });
    const pdfData = await parser.getText();
    await parser.destroy();
    
    const extractedText = pdfData.text || 'No text found in PDF.';
    const doc = new Document({
      sections: [{ properties: {}, children: extractedText.split('\n').map(line => new Paragraph({ children: [new TextRun(line)] })) }]
    });

    const wordBuffer = await Packer.toBuffer(doc);
    const outputFilename = 'converted-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.docx';
    const outputPath = path.join('uploads', outputFilename);
    fs.writeFileSync(outputPath, wordBuffer);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Word)`;
    await saveHistoryRecord(req.user.userId, title, fileUrl, 'docx');
    res.json({ success: true, fileUrl, title });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 3. Word to PDF
router.post('/word-to-pdf', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    const result = await mammoth.extractRawText({ path: req.file.path });
    const extractedText = result.value || 'No text found.';

    const outputFilename = 'converted-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pdf';
    const outputPath = path.join('uploads', outputFilename);
    const doc = new PDFDocument();
    const stream = fs.createWriteStream(outputPath);
    doc.pipe(stream);
    doc.fontSize(16).text(`Converted from: ${req.file.originalname}`, { underline: true });
    doc.moveDown();
    doc.fontSize(12).text(extractedText);
    doc.end();

    stream.on('finish', async () => {
      const fileUrl = `/uploads/${outputFilename}`;
      const title = `${req.file.originalname} (PDF)`;
      await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
      res.json({ success: true, fileUrl, title });
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 4. PDF/Image to PowerPoint (Supports Smart Multi-Slide Chunking & Image Inputs)
router.post('/pdf-to-powerpoint', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const fileExtension = path.extname(req.file.originalname).toLowerCase();
    const isImage = req.file.mimetype.startsWith('image/') || ['.png', '.jpg', '.jpeg'].includes(fileExtension);

    if (isImage) {
      // Handle Image Inputs directly onto a presentation slide
      let slide = pptx.addSlide();
      slide.addText(path.parse(req.file.originalname).name, {
        x: 0.8,
        y: 0.4,
        w: 11.5,
        h: 0.6,
        fontSize: 22,
        bold: true,
        color: '003366'
      });
      slide.addImage({
        path: req.file.path,
        x: 1.5,
        y: 1.2,
        w: 10.0,
        h: 5.0,
        sizing: { type: 'contain', w: 10.0, h: 5.0 }
      });
    } else {
      // Handle PDF Text Parsing and Smart Multi-Slide Chunking
      const dataBuffer = fs.readFileSync(req.file.path);
      const parser = new PDFParse({ data: dataBuffer });
      const pdfData = await parser.getText();
      await parser.destroy();

      const lines = pdfData.text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      
      let slidesData = [];
      let currentTitle = "Overview";
      let currentBullets = [];

      for (let line of lines) {
        const isHeading = line.length < 55 && !line.endsWith('.') && !line.endsWith(',') && (
          line.startsWith('UNIT') || 
          line.includes('Importance of') || 
          line.includes('Aims of') || 
          line.includes('Principles of') || 
          line.includes('Perspective') || 
          line.includes('Overview of') || 
          line.includes('Conceptual model') || 
          line.includes('Things in') || 
          line.includes('Relationships') || 
          line.includes('Diagrams') || 
          line.includes('Rules of') || 
          line.includes('Architecture') || 
          line.includes('Life Cycle') ||
          line === 'Class' || line === 'Interface' || line === 'Collaboration' || line === 'Use case' || line === 'Component' || line === 'Node'
        );

        if (isHeading) {
          if (currentBullets.length > 0 || currentTitle !== "Overview") {
            slidesData.push({ title: currentTitle, bullets: currentBullets });
            currentBullets = [];
          }
          currentTitle = line;
        } else {
          const cleanLine = line.replace(/^[•\-\*]\s*/, '');
          if (cleanLine) {
            currentBullets.push(cleanLine);
          }
        }
      }

      if (currentBullets.length > 0 || currentTitle) {
        slidesData.push({ title: currentTitle, bullets: currentBullets });
      }

      if (slidesData.length === 0) {
        slidesData.push({ title: "Document Content", bullets: [pdfData.text.substring(0, 400)] });
      }

      // Build slides using chunking function to prevent text overcrowding
      slidesData.forEach((slideInfo) => {
        addChunksToSlides(pptx, slideInfo.title, slideInfo.bullets);
      });
    }

    const outputFilename = 'converted-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pptx';
    const outputPath = path.join('uploads', outputFilename);
    await pptx.writeFile({ fileName: outputPath });

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (PowerPoint)`;
    await saveHistoryRecord(req.user.userId, title, fileUrl, 'pptx');
    res.json({ success: true, fileUrl, title });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 5. PowerPoint to PDF
router.post('/powerpoint-to-pdf', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    const fileBuffer = fs.readFileSync(req.file.path);
    const zip = new JSZip();
    const zipContent = await zip.loadAsync(fileBuffer);

    let slideTexts = [];
    let slideIndex = 1;

    while (true) {
      const slideFile = zipContent.file(`ppt/slides/slide${slideIndex}.xml`);
      if (!slideFile) break;

      const slideXml = await slideFile.async('text');
      const matches = slideXml.match(/<a:t[^>]*>(.*?)<\/a:t>/g);
      let textContent = '';
      if (matches) {
        textContent = matches.map(m => m.replace(/<\/?[^>]+(>|$)/g, '')).join(' ');
      }

      slideTexts.push(textContent || `[Slide ${slideIndex} contains graphics or no text]`);
      slideIndex++;
    }

    if (slideTexts.length === 0) {
      slideTexts.push("No readable slides found in presentation.");
    }

    const outputFilename = 'converted-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pdf';
    const outputPath = path.join('uploads', outputFilename);
    const doc = new PDFDocument();
    const stream = fs.createWriteStream(outputPath);
    doc.pipe(stream);

    slideTexts.forEach((slideText, index) => {
      if (index > 0) doc.addPage();
      doc.fontSize(16).text(`Slide ${index + 1}`, { underline: true });
      doc.moveDown();
      doc.fontSize(12).text(slideText, { lineGap: 4 });
    });
    doc.end();

    stream.on('finish', async () => {
      const fileUrl = `/uploads/${outputFilename}`;
      const title = `${req.file.originalname} (PDF)`;
      await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
      res.json({ success: true, fileUrl, title });
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 6. Compress PDF
// 6. Compress PDF (Manual Target Size)


router.post('/compress-pdf', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No PDF uploaded'
      });
    }

    const targetBytes = Number(req.body.targetSizeBytes);

    if (
      !Number.isSafeInteger(targetBytes) ||
      targetBytes < 1024 ||
      targetBytes > 100 * 1024 * 1024
    ) {
      return res.status(400).json({
        success: false,
        error: 'Enter a valid target size between 1 KB and 100 MB.'
      });
    }

    const originalBuffer = fs.readFileSync(req.file.path);
    const originalBytes = originalBuffer.length;

    if (targetBytes >= originalBytes) {
      return res.status(400).json({
        success: false,
        error: 'The compression target must be smaller than the original PDF.'
      });
    }

    const pdfDoc = await PDFLibDocument.load(originalBuffer);

    const compressedBytes = await pdfDoc.save({
      useObjectStreams: true
    });

    if (compressedBytes.length >= originalBytes) {
      return res.status(422).json({
        success: false,
        error: 'This PDF could not be reduced using the current compression method.',
        originalSize: originalBytes,
        achievedSize: compressedBytes.length,
        targetSize: targetBytes
      });
    }

    // Save the reduced PDF even if it misses the requested target.
    const outputFilename =
      'compressed-' + Date.now() + '-' +
      path.parse(req.file.originalname).name + '.pdf';

    const outputPath = path.join('uploads', outputFilename);

    fs.writeFileSync(outputPath, compressedBytes);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Compressed)`;

    await saveHistoryRecord(
      req.user.userId,
      title,
      fileUrl,
      'pdf'
    );

    const achievedTarget = compressedBytes.length <= targetBytes;

    res.json({
      success: true,
      fileUrl,
      title,
      message: achievedTarget
        ? 'PDF compressed and requested target achieved.'
        : 'PDF reduced, but the requested target could not be reached. You can still download the reduced PDF.',
      targetAchieved: achievedTarget,
      ...sizeReport(
        originalBytes,
        compressedBytes.length,
        targetBytes
      )
    });

  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 7. Expand PDF
// 7. Expand PDF (Manual Target Size)

router.post('/expand-pdf', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No PDF uploaded'
      });
    }

    const targetBytes = Number(req.body.targetSizeBytes);

if (
  !Number.isSafeInteger(targetBytes) ||
  targetBytes < 1024 ||
  targetBytes > 100 * 1024 * 1024
) {
  return res.status(400).json({
    success: false,
    error: 'Enter a valid target size between 1 KB and 100 MB.'
  });
}

    const originalBuffer = fs.readFileSync(req.file.path);
    const originalBytes = originalBuffer.length;

    if (targetBytes <= originalBytes) {
      return res.status(400).json({
        success: false,
        error: 'The expansion target must be larger than the original PDF.'
      });
    }

    const pdfDoc = await PDFLibDocument.load(originalBuffer);

    const baseBytes = await pdfDoc.save({
      useObjectStreams: true
    });

    if (baseBytes.length > targetBytes) {
      return res.status(422).json({
        success: false,
        error: 'The processed PDF is already larger than the requested target.'
      });
    }

    const paddingLength = targetBytes - baseBytes.length;
    const padding = Buffer.alloc(paddingLength, 0x20);

    const expandedBytes = Buffer.concat([
      Buffer.from(baseBytes),
      padding
    ]);

    const outputFilename =
      'expanded-' + Date.now() + '-' +
      path.parse(req.file.originalname).name + '.pdf';

    const outputPath = path.join('uploads', outputFilename);

    fs.writeFileSync(outputPath, expandedBytes);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Expanded)`;

    await saveHistoryRecord(
      req.user.userId,
      title,
      fileUrl,
      'pdf'
    );

    res.json({
      success: true,
      fileUrl,
      title,
      ...sizeReport(
        originalBytes,
        expandedBytes.length,
        targetBytes
      )
    });

  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 8. Remove Pages
router.post('/remove-page', verifyToken, upload.single('file'), async (req, res) => {
  try {
    const { pageNumber } = req.body;
    if (!req.file || !pageNumber) return res.status(400).json({ success: false, error: 'File and page numbers required' });

    const pdfDoc = await PDFLibDocument.load(fs.readFileSync(req.file.path));
    const indices = pageNumber.split(',').map(p => parseInt(p.trim()) - 1).filter(p => !isNaN(p));
    indices.sort((a, b) => b - a);

    const total = pdfDoc.getPageCount();
    for (const index of indices) {
      if (index >= 0 && index < total) pdfDoc.removePage(index);
    }

    const modifiedBytes = await pdfDoc.save();
    const outputFilename = 'removed-pages-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pdf';
    const outputPath = path.join('uploads', outputFilename);
    fs.writeFileSync(outputPath, modifiedBytes);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Pages Removed)`;
    await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
    res.json({ success: true, fileUrl, title });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 9. Add Page
router.post('/add-page', verifyToken, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });
    const pdfDoc = await PDFLibDocument.load(fs.readFileSync(req.file.path));
    pdfDoc.addPage([600, 800]);

    const modifiedBytes = await pdfDoc.save();
    const outputFilename = 'added-page-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pdf';
    const outputPath = path.join('uploads', outputFilename);
    fs.writeFileSync(outputPath, modifiedBytes);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Page Added)`;
    await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
    res.json({ success: true, fileUrl, title });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 10. Reorder Pages
router.post('/reorder-pages', verifyToken, upload.single('file'), async (req, res) => {
  try {
    const { pageOrder } = req.body;
    if (!req.file || !pageOrder) return res.status(400).json({ success: false, error: 'File and new page order are required' });

    const existingPdfBytes = fs.readFileSync(req.file.path);
    const pdfDoc = await PDFLibDocument.load(existingPdfBytes);

    const indices = pageOrder.split(',').map(p => parseInt(p.trim()) - 1).filter(p => !isNaN(p));

    if (indices.length === 0) {
      return res.status(400).json({ success: false, error: 'Invalid page order format specified' });
    }

    const newPdfDoc = await PDFLibDocument.create();
    const copiedPages = await newPdfDoc.copyPages(pdfDoc, indices);
    
    copiedPages.forEach((page) => {
      newPdfDoc.addPage(page);
    });

    const modifiedBytes = await newPdfDoc.save();
    const outputFilename = 'reordered-' + Date.now() + '-' + path.parse(req.file.originalname).name + '.pdf';
    const outputPath = path.join('uploads', outputFilename);
    fs.writeFileSync(outputPath, modifiedBytes);

    const fileUrl = `/uploads/${outputFilename}`;
    const title = `${req.file.originalname} (Pages Reordered)`;
    await saveHistoryRecord(req.user.userId, title, fileUrl, 'pdf');
    res.json({ success: true, fileUrl, title });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get User History
router.get('/all', verifyToken, async (req, res) => {
  try {
    const docs = await DocumentModel.find({ userId: req.user.userId }).sort({ createdAt: -1 });
    res.json({ success: true, data: docs });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;