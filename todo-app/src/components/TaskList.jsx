import { useState, useEffect, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";
import {
  getTodoLists,
  getTasksForList,
  createTask,
  updateTask,
  deleteTask,
} from "../graphService";

export default function TaskList() {
  const { instance, accounts } = useMsal();
  const [lists, setLists] = useState([]);
  const [selectedList, setSelectedList] = useState(null);
  const [tasks, setTasks] = useState([]);
  const [newTaskTitle, setNewTaskTitle] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const getToken = useCallback(async () => {
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
    return response.accessToken;
  }, [instance, accounts]);

  const fetchLists = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const token = await getToken();
      const data = await getTodoLists(token);
      setLists(data);
      if (data.length > 0) setSelectedList(data[0]);
    } catch (err) {
      setError("Failed to load task lists.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  }, [getToken]);

  const fetchTasks = useCallback(async (listId) => {
    setLoading(true);
    setError(null);
    try {
      const token = await getToken();
      const data = await getTasksForList(token, listId);
      setTasks(data);
    } catch (err) {
      setError("Failed to load tasks.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  }, [getToken]);

  useEffect(() => { fetchLists(); }, [fetchLists]);
  useEffect(() => { if (selectedList) fetchTasks(selectedList.id); }, [selectedList, fetchTasks]);

  const handleAddTask = async (e) => {
    e.preventDefault();
    if (!newTaskTitle.trim() || !selectedList) return;
    try {
      const token = await getToken();
      await createTask(token, selectedList.id, newTaskTitle.trim());
      setNewTaskTitle("");
      fetchTasks(selectedList.id);
    } catch (err) {
      setError("Failed to create task.");
    }
  };

  const handleToggleTask = async (task) => {
    try {
      const token = await getToken();
      const newStatus = task.status === "completed" ? "notStarted" : "completed";
      await updateTask(token, selectedList.id, task.id, { status: newStatus });
      fetchTasks(selectedList.id);
    } catch (err) {
      setError("Failed to update task.");
    }
  };

  const handleDeleteTask = async (taskId) => {
    try {
      const token = await getToken();
      await deleteTask(token, selectedList.id, taskId);
      fetchTasks(selectedList.id);
    } catch (err) {
      setError("Failed to delete task.");
    }
  };

  const handleLogout = () => {
    instance.logoutPopup().catch(console.error);
  };

  const pendingTasks = tasks.filter((t) => t.status !== "completed");
  const completedTasks = tasks.filter((t) => t.status === "completed");

  return (
    <div className="app-container">
      <aside className="sidebar">
        <div className="sidebar-header">
          <span className="sidebar-logo">✔</span>
          <span>To-Do</span>
        </div>
        <nav className="list-nav">
          {lists.map((list) => (
            <button
              key={list.id}
              className={`list-item ${selectedList?.id === list.id ? "active" : ""}`}
              onClick={() => setSelectedList(list)}
            >
              <span className="list-icon">📋</span>
              {list.displayName}
            </button>
          ))}
        </nav>
        <button className="btn btn-logout" onClick={handleLogout}>
          Sign out
        </button>
      </aside>

      <main className="main-content">
        {error && <div className="error-banner">{error}</div>}

        <h2 className="list-title">{selectedList?.displayName ?? "Select a list"}</h2>

        <form className="add-task-form" onSubmit={handleAddTask}>
          <input
            type="text"
            className="task-input"
            placeholder="Add a task"
            value={newTaskTitle}
            onChange={(e) => setNewTaskTitle(e.target.value)}
          />
          <button type="submit" className="btn btn-primary" disabled={!newTaskTitle.trim()}>
            Add
          </button>
        </form>

        {loading ? (
          <div className="loading">Loading...</div>
        ) : (
          <>
            <ul className="task-list">
              {pendingTasks.map((task) => (
                <TaskItem
                  key={task.id}
                  task={task}
                  onToggle={handleToggleTask}
                  onDelete={handleDeleteTask}
                />
              ))}
            </ul>

            {completedTasks.length > 0 && (
              <>
                <h3 className="completed-header">Completed ({completedTasks.length})</h3>
                <ul className="task-list task-list--completed">
                  {completedTasks.map((task) => (
                    <TaskItem
                      key={task.id}
                      task={task}
                      onToggle={handleToggleTask}
                      onDelete={handleDeleteTask}
                    />
                  ))}
                </ul>
              </>
            )}

            {tasks.length === 0 && !loading && (
              <p className="empty-state">No tasks yet. Add one above!</p>
            )}
          </>
        )}
      </main>
    </div>
  );
}

function TaskItem({ task, onToggle, onDelete }) {
  const completed = task.status === "completed";
  return (
    <li className={`task-item ${completed ? "task-item--completed" : ""}`}>
      <button
        className={`task-checkbox ${completed ? "checked" : ""}`}
        onClick={() => onToggle(task)}
        aria-label={completed ? "Mark incomplete" : "Mark complete"}
      >
        {completed && "✓"}
      </button>
      <span className="task-title">{task.title}</span>
      {task.dueDateTime && (
        <span className="task-due">
          {new Date(task.dueDateTime.dateTime).toLocaleDateString()}
        </span>
      )}
      <button
        className="task-delete"
        onClick={() => onDelete(task.id)}
        aria-label="Delete task"
      >
        ✕
      </button>
    </li>
  );
}
