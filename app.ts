import express, { Request, Response, RequestHandler } from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { TemplateHandler, MimeType } from 'easy-template-x';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as ExcelJS from 'exceljs';
import { downloadHandler } from './routes/download';
import { supabase } from './supabaseClient';

const execAsync = promisify(exec);
const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.set('view engine', 'ejs');

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req: Express.Request, file: Express.Multer.File, cb: (error: Error | null, destination: string) => void) => {
    cb(null, path.join(__dirname, 'templates'));
  },
  filename: (req: Express.Request, file: Express.Multer.File, cb: (error: Error | null, filename: string) => void) => {
    cb(null, file.originalname);
  }
});

const upload = multer({ storage });

// Directory paths
const DIRS = {
  assets: path.resolve(__dirname, 'assets'),
  templates: path.resolve(__dirname, 'templates'),
  generatedDocx: path.resolve(__dirname, 'output-generated', 'docx'),
  generatedPdf: path.resolve(__dirname, 'output-generated', 'pdf'),
  generatedExcel: path.resolve(__dirname, 'output-generated', 'excel')
};

// Ensure directories exist
Object.values(DIRS).forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Function to find placeholders in Excel cells
async function findExcelPlaceholders(filePath: string): Promise<Set<string>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const placeholders = new Set<string>();
  const placeholderRegex = /{{([^}]+)}}/g;

  console.log('Searching for placeholders in Excel file:', filePath);
  console.log('Number of worksheets:', workbook.worksheets.length);
  
  workbook.worksheets.forEach((worksheet: ExcelJS.Worksheet) => {
    console.log('Checking worksheet:', worksheet.name);
    let rowCount = 0;
    worksheet.eachRow((row: ExcelJS.Row) => {
      rowCount++;
      row.eachCell((cell: ExcelJS.Cell) => {
        const rawValue = cell.value;
        console.log(`Row ${rowCount}, Column ${cell.col}, Raw value:`, rawValue);
        
        let cellValue = '';
        if (typeof rawValue === 'string') {
          cellValue = rawValue;
        } else if (rawValue && typeof rawValue === 'object' && 'result' in rawValue) {
          // Handle formula cells
          cellValue = rawValue.result?.toString() || '';
        } else if (rawValue && typeof rawValue === 'object' && 'text' in rawValue) {
          // Handle rich text cells
          cellValue = rawValue.text || '';
        }
        
        if (cellValue) {
          const matches = cellValue.match(placeholderRegex);
          if (matches) {
            matches.forEach(match => {
              // Remove {{ and }} to get the placeholder name
              const placeholder = match.slice(2, -2);
              console.log('Found placeholder:', placeholder, 'in cell value:', cellValue);
              placeholders.add(placeholder);
            });
          }
        }
      });
    });
    console.log(`Processed ${rowCount} rows in worksheet ${worksheet.name}`);
  });

  console.log('Total placeholders found:', placeholders.size);
  return placeholders;
}

// Function to prepare Excel file for PDF conversion - now just preserves the original format
async function prepareExcelForPdf(workbook: ExcelJS.Workbook): Promise<void> {
  // No modifications to preserve original Excel structure
  return;
}

// Function to replace placeholders in Excel file
async function processExcelTemplate(templatePath: string, outputPath: string, formData: Record<string, any>): Promise<void> {
  // First load the template to capture all original formatting
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.readFile(templatePath);
  
  // Create a new workbook for processing while preserving all formatting
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const placeholderRegex = /{{([^}]+)}}/g;

  workbook.worksheets.forEach((worksheet: ExcelJS.Worksheet, sheetIndex: number) => {
    const templateSheet = templateWorkbook.worksheets[sheetIndex];

    // Preserve exact column widths
    worksheet.columns = templateSheet.columns.map((col: Partial<ExcelJS.Column>) => ({
      ...col,
      width: col.width // Explicitly preserve width
    }));

    worksheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
      const templateRow = templateSheet.getRow(rowNumber);
      
      // Preserve row dimensions and properties
      row.height = templateRow.height;
      row.hidden = templateRow.hidden;
      row.outlineLevel = templateRow.outlineLevel;
      
      row.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
        const templateCell = templateRow.getCell(colNumber);
        
        // Keep the original style reference
        const originalStyle = JSON.parse(JSON.stringify(templateCell.style));
        
        if (typeof cell.value === 'string') {
          const cellValue = cell.value;
          if (cellValue.match(placeholderRegex)) {
            // Replace placeholders while preserving data types
            const newValue = cellValue.replace(placeholderRegex, (match: string, placeholder: string) => {
              const value = formData[placeholder];
              
              if (value === undefined) return match;
              
              // Preserve data types
              if (typeof value === 'string') {
                // Handle dates
                if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
                  return new Date(value);
                }
                // Handle numbers
                if (!isNaN(Number(value)) && value.trim() !== '') {
                  return Number(value);
                }
              }
              
              return value;
            });
            
            cell.value = newValue;
          }
        }
        
        // Always restore the exact original style
        cell.style = originalStyle;
      });
    });

    // Preserve worksheet-level properties
    worksheet.properties = JSON.parse(JSON.stringify(templateSheet.properties));
    worksheet.pageSetup = JSON.parse(JSON.stringify(templateSheet.pageSetup));
    
    // Preserve column width settings explicitly
    templateSheet.columns.forEach((col: Partial<ExcelJS.Column>, index: number) => {
      if (col && col.width) {
        worksheet.getColumn(index + 1).width = col.width;
      }
    });
  });

  // Set workbook properties and views from template
  workbook.properties = JSON.parse(JSON.stringify(templateWorkbook.properties));
  workbook.views = JSON.parse(JSON.stringify(templateWorkbook.views));

  await workbook.xlsx.writeFile(outputPath);
}

// Routes
app.get('/', (req: Request, res: Response) => {
  res.render('index');
});

// Upload template and extract placeholders
app.post('/upload', upload.single('template'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      throw new Error('No file uploaded');
    }

    // Upload file to Supabase Storage
    const filePath = path.join(DIRS.templates, req.file.originalname);
    const fileBuffer = fs.readFileSync(filePath);
    const { data: uploadData, error: uploadError } = await supabase.storage
      .from('uploads') // Change 'uploads' to your bucket name if different
      .upload(req.file.originalname, fileBuffer, {
        contentType: req.file.mimetype,
        upsert: true,
      });
    if (uploadError) {
      console.error('Supabase upload error:', uploadError);
      throw new Error('Failed to upload file to Supabase Storage: ' + uploadError.message);
    }
    console.log('File uploaded to Supabase Storage:', uploadData);

    const templatePath = path.join(DIRS.templates, req.file.originalname);
    const fileExt = path.extname(req.file.originalname).toLowerCase();
    let uniquePlaceholders: Set<string>;
    if (fileExt === '.xlsx') {
      // Handle Excel template
      console.log('Processing Excel template:', templatePath);
      uniquePlaceholders = await findExcelPlaceholders(templatePath);
      console.log('Placeholders found in Excel:', Array.from(uniquePlaceholders));
    } else if (fileExt === '.docx') {
      // Handle Word template
      const templateBuffer = await fs.promises.readFile(templatePath);
      const handler = new TemplateHandler();
      const tags = await handler.parseTags(templateBuffer);
      uniquePlaceholders = new Set<string>();
      for (const tag of tags) {
        console.log('Found placeholder:', tag.name);
        uniquePlaceholders.add(tag.name);
      }
    } else {
      throw new Error('Unsupported file format. Please upload a .docx or .xlsx file.');
    }

    if (uniquePlaceholders.size === 0) {
      throw new Error('No placeholders found in template. Make sure your template has placeholders in the format {{placeholder}}');
    }

    console.log(`Found ${uniquePlaceholders.size} unique placeholders`);
    
    return res.render('form', { 
      placeholders: Array.from(uniquePlaceholders),
      templateName: req.file.originalname
    });

  } catch (error: any) {
    console.error('Error processing template:', error);
    res.status(500).send(error.message);
  }
});

// Generate document from form data
app.post('/generate', async (req, res) => {
  try {
    const { templateName, formData } = req.body;
    const fileExt = path.extname(templateName).toLowerCase();
    const templatePath = path.join(DIRS.templates, templateName);
    const timestamp = new Date().getTime();
      if (fileExt === '.xlsx') {
      // Handle Excel template
      const outputExcel = path.join(DIRS.generatedExcel, `generated-${timestamp}.xlsx`);
      const outputPdf = path.join(DIRS.generatedPdf, `generated-${timestamp}.pdf`);
      
      // Process template with form data
      await processExcelTemplate(templatePath, outputExcel, formData);
      
      // Convert to PDF
      await convertExcelToPdf(outputExcel, outputPdf);
      
      res.json({ 
        message: 'Files generated successfully',
        excelFilename: `generated-${timestamp}.xlsx`,
        pdfFilename: `generated-${timestamp}.pdf`,
        fileType: 'excel'
      });
    } else if (fileExt === '.docx') {
      // Handle Word template
      // Convert base64 image data to buffers
      Object.entries(formData).forEach(([key, value]: [string, any]) => {
        if (value && value._type === 'image') {
          console.log(`Processing image for ${key}`);
          // Convert base64 to Buffer
          const imageBuffer = Buffer.from(value.source, 'base64');
          // Update the image data to match the easy-template-x format
          formData[key] = {
            _type: 'image',
            source: imageBuffer,
            format: MimeType.Png,
            width: 150,
            height: 100,
            altText: value.altText || key,
            transparencyPercent: value.transparencyPercent || 0
          };

          console.log(`Image processed: ${key}`, {
            format: 'png',
            size: imageBuffer.length,
            width: value.width || 200,
            height: value.height || 200
          });
        }
      });

      const docxFilename = `generated-${timestamp}.docx`;
      const pdfFilename = `generated-${timestamp}.pdf`;
      const outputDocx = path.join(DIRS.generatedDocx, docxFilename);
      const outputPdf = path.join(DIRS.generatedPdf, pdfFilename);

      // Read template as buffer
      const templateContent = await fs.promises.readFile(templatePath);

      // Process template with form data
      const handler = new TemplateHandler();
      const doc = await handler.process(templateContent, formData);

      // Save generated DOCX
      await fs.promises.writeFile(outputDocx, doc);      
      // Convert to PDF
      await convertToPdf(outputDocx, outputPdf);
      
      // Debug log
      console.log('Generated files:', {
        docxFilename,
        pdfFilename,
        fileType: 'docx'
      });
      
      res.json({ 
        message: 'Files generated successfully',
        docxFilename,
        pdfFilename,
        fileType: 'docx'
      });
    } else {
      throw new Error('Unsupported file format');
    }
  } catch (error: any) {
    console.error('Error processing template:', error);
    res.status(500).send(error.message);
  }
});

// Download route handler
app.get('/download/:type/:filename', downloadHandler);

// Function to convert DOCX to PDF using LibreOffice
async function convertToPdf(inputPath: string, outputPath: string): Promise<void> {
  try {
    const absoluteInputPath = path.resolve(inputPath);
    const absoluteOutputDir = path.resolve(path.dirname(outputPath));
    
    if (!fs.existsSync(absoluteInputPath)) {
      throw new Error(`Input file not found: ${absoluteInputPath}`);
    }
    
    if (!fs.existsSync(absoluteOutputDir)) {
      fs.mkdirSync(absoluteOutputDir, { recursive: true });
    }

    const command = `soffice --headless --norestore --convert-to pdf:writer_pdf_Export --outdir "${absoluteOutputDir}" "${absoluteInputPath}"`;
    const { stdout, stderr } = await execAsync(command);
    
    const expectedPdfPath = path.join(absoluteOutputDir, path.basename(absoluteInputPath, '.docx') + '.pdf');
    if (!fs.existsSync(expectedPdfPath)) {
      throw new Error('PDF file was not created after conversion');
    }
  } catch (error: any) {
    throw new Error(`PDF conversion error: ${error.message}`);
  }
}

// Function to convert Excel to PDF using LibreOffice
async function convertExcelToPdf(inputPath: string, outputPath: string): Promise<void> {
  try {
    const absoluteInputPath = path.resolve(inputPath);
    const absoluteOutputDir = path.resolve(path.dirname(outputPath));
    
    console.log('Converting Excel to PDF:');
    console.log('Input path:', absoluteInputPath);
    console.log('Output directory:', absoluteOutputDir);
    
    if (!fs.existsSync(absoluteInputPath)) {
      throw new Error(`Input file not found: ${absoluteInputPath}`);
    }
    
    if (!fs.existsSync(absoluteOutputDir)) {
      fs.mkdirSync(absoluteOutputDir, { recursive: true });
    }

    // Construct a simpler, more reliable command
    const command = `soffice --headless --convert-to pdf --outdir "${absoluteOutputDir}" "${absoluteInputPath}"`;
    console.log('Executing command:', command);

    const { stdout, stderr } = await execAsync(command);
    console.log('Conversion output:', stdout);
    if (stderr) {
      console.error('Conversion errors:', stderr);
    }
    
    const expectedPdfPath = path.join(absoluteOutputDir, path.basename(absoluteInputPath, '.xlsx') + '.pdf');
    console.log('Expected PDF path:', expectedPdfPath);

    // Wait a short time to ensure file system has completed writing
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    if (!fs.existsSync(expectedPdfPath)) {
      throw new Error(`PDF file was not created at expected path: ${expectedPdfPath}`);
    }

    console.log('PDF conversion completed successfully');
  } catch (error: any) {
    console.error('Detailed conversion error:', error);
    throw new Error(`Excel to PDF conversion error: ${error.message}`);
  }
}

// Function to cleanup old generated files
function cleanupGeneratedFiles(maxAgeHours: number = 1): void {
  const now = Date.now();
  const maxAge = maxAgeHours * 60 * 60 * 1000; // Convert hours to milliseconds

  // Clean each output directory
  ['generatedDocx', 'generatedPdf', 'generatedExcel'].forEach(dirKey => {
    const dir = DIRS[dirKey as keyof typeof DIRS];
    if (fs.existsSync(dir)) {
      fs.readdirSync(dir).forEach(file => {
        const filePath = path.join(dir, file);
        const stats = fs.statSync(filePath);
        
        if (now - stats.mtimeMs > maxAge) {
          try {
            fs.unlinkSync(filePath);
            console.log(`Cleaned up old file: ${filePath}`);
          } catch (error) {
            console.error(`Error cleaning up file ${filePath}:`, error);
          }
        }
      });
    }
  });
}

// Run cleanup on server start and every 12 hours
cleanupGeneratedFiles();
setInterval(() => cleanupGeneratedFiles(), 12 * 60 * 60 * 1000);

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

async function fetchUsers() {
  const { data, error } = await supabase.from('users').select('*');
  if (error) {
    console.error('Error fetching users:', error.message);
  } else {
    console.log('Fetched users:', data);
  }
}

fetchUsers();
