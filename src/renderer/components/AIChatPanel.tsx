import { useState } from "react";

const AIChatPanel = () => {
  const [input, setInput] = useState("");
  const [messages, setMessages] = useState<string[]>([]);

  const handleSend = () => {
    if (!input.trim()) return;

    setMessages((prev) => [...prev, input]);
    setInput("");
  };

  return (
    <aside className="w-80 bg-white dark:bg-gray-900 border-l border-gray-200 dark:border-gray-700 p-4 flex flex-col">
      
      {/* Messages */}
      <div className="flex-1 overflow-auto mb-4">
        {messages.length === 0 ? (
          <p className="text-gray-400 text-sm">No messages yet</p>
        ) : (
          messages.map((msg, i) => (
            <div
              key={i}
              className="mb-2 p-2 rounded bg-gray-100 dark:bg-gray-800 text-gray-900 dark:text-gray-100"
            >
              {msg}
            </div>
          ))
        )}
      </div>

      {/* Input */}
      <div className="flex gap-2">
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Type a message..."
          className="flex-1 border rounded px-2 py-1 bg-white dark:bg-gray-800 text-gray-900 dark:text-white"
        />

        <button
          onClick={handleSend}
          className="px-3 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded"
        >
          Send
        </button>
      </div>

    </aside>
  );
};

export default AIChatPanel;