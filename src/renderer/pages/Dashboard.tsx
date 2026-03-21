import Sidebar from "../components/Sidebar";
import Header from "../components/Header";
import AIChatPanel from "../components/AIChatPanel";

const Dashboard = () => {
  return (
    <div className="flex h-screen bg-[#f6f8fb] dark:bg-[#0b0f19]">
      <Sidebar />

      <div className="flex flex-col flex-1">
        <Header />

        <main className="flex-1 p-6 overflow-auto">
          <h1 className="text-2xl font-semibold text-gray-900 dark:text-white mb-6">
            Dashboard
          </h1>

          {/* Cards Grid */}
          <div className="grid grid-cols-3 gap-6">
            
            <div className="p-5 rounded-xl bg-white dark:bg-[#111827] shadow-sm hover:shadow-md transition">
              <h2 className="text-lg font-medium mb-2">Excel Generator</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">
                Generate Excel formulas using AI prompts.
              </p>
            </div>

            <div className="p-5 rounded-xl bg-white dark:bg-[#111827] shadow-sm hover:shadow-md transition">
              <h2 className="text-lg font-medium mb-2">Recent Files</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">
                View and manage your recent work.
              </p>
            </div>

            <div className="p-5 rounded-xl bg-white dark:bg-[#111827] shadow-sm hover:shadow-md transition">
              <h2 className="text-lg font-medium mb-2">AI Assistant</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">
                Ask questions and automate workflows.
              </p>
            </div>

          </div>
        </main>
      </div>

      <AIChatPanel />
    </div>
  );
};

export default Dashboard;