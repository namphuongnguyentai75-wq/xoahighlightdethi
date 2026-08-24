import { useState, useRef } from 'react';
import axios from 'axios';
import { UploadCloud, File, CheckCircle, Loader2, Download, AlertCircle } from 'lucide-react';

function App() {
  const [file, setFile] = useState(null);
  const [outputFormat, setOutputFormat] = useState('docx');

  const [isProcessing, setIsProcessing] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [error, setError] = useState(null);
  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const droppedFile = e.dataTransfer.files[0];
    validateAndSetFile(droppedFile);
  };

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    validateAndSetFile(selectedFile);
  };

  const validateAndSetFile = (selectedFile) => {
    if (!selectedFile) return;
    
    const validTypes = ['application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
    if (!validTypes.includes(selectedFile.type) && !selectedFile.name.match(/\.(pdf|docx)$/i)) {
      setError("Chỉ hỗ trợ định dạng .docx và .pdf");
      return;
    }
    
    setError(null);
    setFile(selectedFile);
    setDownloadUrl(null);
  };



  const handleProcess = async () => {
    if (!file) return;

    setIsProcessing(true);
    setError(null);
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('output_format', outputFormat);

    try {
      const response = await axios.post('http://localhost:8000/api/process', formData, {
        responseType: 'blob', // Important for downloading files
      });

      // Get filename from response header or default
      const contentDisposition = response.headers['content-disposition'];
      let filename = `cleaned_${file.name}`;
      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename="?(.+)"?/);
        if (filenameMatch && filenameMatch.length === 2)
          filename = filenameMatch[1];
      }
      
      // If output format changed to pdf
      if (outputFormat === 'pdf' && !filename.endsWith('.pdf')) {
          filename = filename.replace(/\.docx$/i, '.pdf');
      }

      // Create download link
      const url = window.URL.createObjectURL(new Blob([response.data]));
      setDownloadUrl({ url, filename });

    } catch (err) {
      console.error("Processing error:", err);
      setError("Đã xảy ra lỗi trong quá trình xử lý. Vui lòng thử lại.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col items-center py-12 px-4 sm:px-6 lg:px-8">
      <div className="w-full max-w-3xl space-y-8">
        
        {/* Header */}
        <div className="text-center">
          <h2 className="mt-6 text-3xl font-extrabold text-gray-900">
            Trình Phục Chế Form Đề Thi
          </h2>
          <p className="mt-2 text-sm text-gray-600">
            Tự động bóc tách và ép lại chuẩn form câu hỏi, đáp án, và hình ảnh. Mọi vết highlight đều biến mất hoàn toàn.
          </p>
        </div>

        <div className="bg-white shadow rounded-lg p-6 sm:p-8">
          
          {/* Upload Zone */}
          <div 
            className={`mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-gray-300 border-dashed rounded-md transition-colors ${file ? 'bg-blue-50 border-blue-400' : 'hover:border-gray-400'}`}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            onClick={() => !file && fileInputRef.current.click()}
          >
            <div className="space-y-1 text-center">
              {file ? (
                <div className="flex flex-col items-center">
                  <File className="mx-auto h-12 w-12 text-blue-500" />
                  <p className="mt-2 text-sm text-gray-700 font-medium">{file.name}</p>
                  <p className="text-xs text-gray-500">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                  <button 
                    className="mt-4 text-sm text-red-600 hover:text-red-800"
                    onClick={(e) => { e.stopPropagation(); setFile(null); setDownloadUrl(null); }}
                  >
                    Xóa file
                  </button>
                </div>
              ) : (
                <>
                  <UploadCloud className="mx-auto h-12 w-12 text-gray-400" />
                  <div className="flex text-sm text-gray-600 justify-center">
                    <label
                      htmlFor="file-upload"
                      className="relative cursor-pointer bg-white rounded-md font-medium text-blue-600 hover:text-blue-500 focus-within:outline-none focus-within:ring-2 focus-within:ring-offset-2 focus-within:ring-blue-500"
                    >
                      <span>Tải file lên</span>
                      <input id="file-upload" name="file-upload" type="file" className="sr-only" ref={fileInputRef} onChange={handleFileChange} accept=".docx,.pdf" />
                    </label>
                    <p className="pl-1">hoặc kéo thả vào đây</p>
                  </div>
                  <p className="text-xs text-gray-500">
                    Chấp nhận DOCX, PDF (Tối đa 50MB)
                  </p>
                </>
              )}
            </div>
          </div>

          {error && (
            <div className="mt-4 bg-red-50 border-l-4 border-red-400 p-4">
              <div className="flex">
                <div className="flex-shrink-0">
                  <AlertCircle className="h-5 w-5 text-red-400" aria-hidden="true" />
                </div>
                <div className="ml-3">
                  <p className="text-sm text-red-700">{error}</p>
                </div>
              </div>
            </div>
          )}

          {/* Options Panel */}
          <div className="mt-8 grid grid-cols-1 gap-8 border-t border-gray-200 pt-8">
            
            <div className="flex justify-center">
              <div className="space-y-4">
                <h3 className="text-lg font-medium text-gray-900 mb-4 text-center">Tùy chọn xuất file</h3>
                <div className="flex space-x-6 justify-center">
                  <div className="flex items-center">
                    <input
                      id="format-docx"
                      name="outputFormat"
                      type="radio"
                      value="docx"
                      checked={outputFormat === 'docx'}
                      onChange={(e) => setOutputFormat(e.target.value)}
                      className="focus:ring-blue-500 h-4 w-4 text-blue-600 border-gray-300"
                    />
                    <label htmlFor="format-docx" className="ml-3 block text-sm font-medium text-gray-700">
                      Xuất file Word (.docx)
                    </label>
                  </div>
                  <div className="flex items-center">
                    <input
                      id="format-pdf"
                      name="outputFormat"
                      type="radio"
                      value="pdf"
                      checked={outputFormat === 'pdf'}
                      onChange={(e) => setOutputFormat(e.target.value)}
                      className="focus:ring-blue-500 h-4 w-4 text-blue-600 border-gray-300"
                    />
                    <label htmlFor="format-pdf" className="ml-3 block text-sm font-medium text-gray-700">
                      Xuất file PDF (.pdf)
                    </label>
                  </div>
                </div>
              </div>
            </div>

          </div>

          {/* Actions */}
          <div className="mt-8 pt-5 border-t border-gray-200 flex justify-center">
            {!downloadUrl ? (
              <button
                type="button"
                onClick={handleProcess}
                disabled={!file || isProcessing}
                className={`inline-flex items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white ${!file || isProcessing ? 'bg-blue-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500'}`}
              >
                {isProcessing ? (
                  <>
                    <Loader2 className="animate-spin -ml-1 mr-3 h-5 w-5 text-white" />
                    Đang xử lý...
                  </>
                ) : (
                  <>
                    Bắt đầu làm sạch
                  </>
                )}
              </button>
            ) : (
              <a
                href={downloadUrl.url}
                download={downloadUrl.filename}
                className="inline-flex items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
              >
                <Download className="-ml-1 mr-3 h-5 w-5" />
                Tải xuống tài liệu ({downloadUrl.filename})
              </a>
            )}
          </div>

        </div>
      </div>
    </div>
  );
}

export default App;
