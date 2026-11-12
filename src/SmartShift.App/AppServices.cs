using SmartShift.Core.Actions;
using SmartShift.Core.Capabilities;
using SmartShift.Core.Classification;
using SmartShift.Core.Logging;
using SmartShift.Core.Queueing;
using SmartShift.Infrastructure.Logging;
using SmartShift.Infrastructure.Output;
using SmartShift.Processing;
using SmartShift.Processing.Batch;
using SmartShift.Processing.Conversion;
using SmartShift.Processing.Documents;
using SmartShift.Processing.Documents.Editing;
using SmartShift.Processing.History;
using SmartShift.Processing.Images;
using SmartShift.Processing.Pdf;
using SmartShift.Processing.Pdf.Editing;
using SmartShift.Processing.PresentationStudio.Service;
using SmartShift.Processing.Quality;
using SmartShift.Processing.Workflows;

namespace SmartShift.App;

public sealed class AppServices : IDisposable
{
    public AppServices(DiagnosticLogger logger)
    {
        Logger = logger;
        Classifier = new FileClassifier();
        Capabilities = new FileCapabilityRegistry();
        Queue = new FileQueue();
        Recommendations = new RecommendationEngine(new UniversalActionCapabilities());
        Outputs = new OutputPathAllocator();
        PdfConversion = new PdfConversionEngine(Outputs);
        OfficeConversion = new OfficeConversionEngine(Outputs);
        Processors = new ProcessorCatalog(Classifier, Outputs);
        PdfMerge = new PdfMergeEngine(Outputs);
        PdfSplit = new PdfSplitEngine(Outputs);
        PdfOrganize = new PdfPageOrganizerEngine(Outputs);
        PdfWatermark = new PdfWatermarkService(Outputs);
        PdfMetadata = new PdfMetadataService(Outputs);
        DocxMerge = new DocxMergeService(Outputs);
        DocxPageNumber = new DocxPageNumberService(Outputs);
        DocxWatermark = new DocxWatermarkService(Outputs);
        DocxMetadata = new DocxMetadataService(Outputs);
        QualityAnalyzer = new ConversionQualityAnalyzer();
        UnifiedConversion = new UnifiedConversionService(Capabilities, Processors, Outputs, Classifier);
        Processors.Register(ActionIds.UnifiedConvert, new UnifiedConvertProcessor(UnifiedConversion));
        BatchEngine = new BatchImageProcessingEngine(Processors, Outputs, Classifier);
        DocumentBuilder = new MultiImageDocumentBuilder(Outputs, Classifier);
        PresentationStudio = new PresentationStudioService(Outputs);
        ImagePresets = new ImagePresetStore();
        HistoryStore = new ProcessingHistoryStore();
        WorkflowStore = new SavedWorkflowStore();
        RetryService = new OperationRetryService(Processors, Outputs, Classifier, HistoryStore);
        WorkflowExecution = new WorkflowExecutionService(Processors, Outputs, HistoryStore);
        LogSink = new LocalFileLogSink(logger);
    }

    public DiagnosticLogger Logger { get; }
    public FileClassifier Classifier { get; }
    public IFileCapabilityRegistry Capabilities { get; }
    public FileQueue Queue { get; }
    public RecommendationEngine Recommendations { get; }
    public ProcessorCatalog Processors { get; }
    public IPdfMergeService PdfMerge { get; }
    public IPdfSplitService PdfSplit { get; }
    public IPdfPageOrganizerService PdfOrganize { get; }
    public IPdfWatermarkService PdfWatermark { get; }
    public IPdfMetadataService PdfMetadata { get; }
    public IDocxMergeService DocxMerge { get; }
    public IDocxPageNumberService DocxPageNumber { get; }
    public IDocxWatermarkService DocxWatermark { get; }
    public IDocxMetadataService DocxMetadata { get; }
    public IConversionQualityAnalyzer QualityAnalyzer { get; }
    public IPdfConversionEngine PdfConversion { get; }
    public IOfficeConversionEngine OfficeConversion { get; }
    public IUnifiedConversionService UnifiedConversion { get; }
    public IBatchImageProcessingEngine BatchEngine { get; }
    public IMultiImageDocumentBuilder DocumentBuilder { get; }
    public IPresentationStudioService PresentationStudio { get; }
    public ImagePresetStore ImagePresets { get; }
    public IProcessingHistoryStore HistoryStore { get; }
    public ISavedWorkflowStore WorkflowStore { get; }
    public IOperationRetryService RetryService { get; }
    public IWorkflowExecutionService WorkflowExecution { get; }
    public OutputPathAllocator Outputs { get; }
    public LocalFileLogSink LogSink { get; }

    public void Dispose() => LogSink.Dispose();
}
