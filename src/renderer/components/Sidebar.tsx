import { NavLink } from "react-router-dom";

const Sidebar = () => {
  const links = [
    "Dashboard",
    "Excel Sheet",
    "Documents",
    "File Manager",
    "Workflow",
    "Developer Mode",
    "Settings",
  ];

  return (
    <aside className="w-56 bg-white dark:bg-gray-900 border-r border-gray-200 dark:border-gray-700 p-3 space-y-2">
      {links.map((link) => (
        <NavLink
          key={link}
          to="/"
          className={({ isActive }) =>
            `block px-3 py-2 rounded text-gray-800 dark:text-gray-100 transition ${
              isActive
                ? "bg-indigo-600 text-white"
                : "hover:bg-gray-100 dark:hover:bg-gray-800"
            }`
          }
        >
          {link}
        </NavLink>
      ))}
    </aside>
  );
};

export default Sidebar;