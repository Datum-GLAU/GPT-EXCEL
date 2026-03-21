const AIChatPanel = () => {
  return (
    <div className="w-80 flex flex-col bg-white dark:bg-gray-900 border-l border-gray-200 dark:border-gray-700 p-4">
      <div className="flex-1 overflow-auto text-gray-800 dark:text-gray-100 mb-4">
        Chat messages
      </div>

      <div className="flex gap-2">
        <input
          type="text"
          placeholder="Type a message..."
          className="flex-1 border rounded px-2 py-1 bg-white dark:bg-gray-800 text-gray-800 dark:text-gray-100 placeholder-gray-500"
        />

        <button className="px-3 py-1 rounded bg-indigo-600 hover:bg-indigo-700 text-white transition">
          Send
        </button>
      </div>
    </div>
  );
};

export default AIChatPanel;