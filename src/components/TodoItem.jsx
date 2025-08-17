import React from 'react';

const TodoItem = ({ todo, onToggle, onDelete }) => {
  return (
    <div className={`todo-item ${todo.completed ? 'completed' : ''}`}>
      <input
        type="checkbox"
        checked={todo.completed}
        onChange={() => onToggle(todo.id)}
      />
      <span className="todo-text">{todo.text}</span>
      <button
        className="delete-btn"
        onClick={() => onDelete(todo.id)}
      >
        Delete
      </button>
    </div>
  );
};

export default TodoItem;