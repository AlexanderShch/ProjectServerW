#pragma once

namespace ProjectServerW {

	// Экспорт в Excel вынесен из DataForm, чтобы UI-логика и COM-логика не жили в одном файле.
	public ref class FormExcel abstract sealed
	{
	public:
		ref class ExcelExportJob sealed {
		public:
			System::Data::DataTable^ tableSnapshot;
			System::String^ saveDirectory;
			System::DateTime sessionStart;
			System::DateTime sessionEnd;
			int clientPort;
			System::String^ clientIP;
			System::String^ formGuid;
			System::WeakReference^ formRef;
			bool enableButtonOnComplete;
		};

		static void EnqueueExport(ExcelExportJob^ job);
		static System::Threading::Mutex^ GetExcelGlobalMutex();
		static bool WaitForAllExports(int timeoutMs);

	private:
		static System::Threading::Mutex^ excelGlobalMutex;
		static System::Collections::Concurrent::ConcurrentQueue<ExcelExportJob^>^ excelExportQueue;
		static System::Threading::AutoResetEvent^ excelExportQueueEvent;
		static System::Threading::Thread^ excelExportWorkerThread;
		static System::Object^ excelExportWorkerSync;
		static bool excelExportWorkerStarted;
		static System::Threading::ManualResetEvent^ excelAllExportsCompletedEvent;
		static int excelActiveExportJobs;

		static FormExcel();
		static void EnsureExcelExportWorker();
		static void ExcelExportWorkerLoop();
		static void ProcessExcelExportJob(ExcelExportJob^ job);
		static System::String^ ResolveExcelSaveDirectory(System::String^ preferredDirectory);
		static void UpdateExcelIdleState();
	};

} // namespace ProjectServerW


