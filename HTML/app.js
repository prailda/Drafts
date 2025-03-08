// Структура данных приложения
let state = {
    projects: [],
    tasks: [],
    categories: [],
    currentProject: null,
    views: {
        current: 'tree', // tree, kanban, timeline, stats
        filters: {
            category: null,
            priority: null,
            status: null,
            search: ''
        }
    }
};

function loadCategories() {
    return fetchCategoriesFromAPI().then(categories => {
        state.categories = categories;
        updateCategorySelect();
    }).catch(error => {
        console.error('Error loading categories:', error);
    });
}

// Вспомогательные функции
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2);
}

function saveState() {
    try {
        const serializedState = JSON.stringify(state);
        if (!serializedState) {
            throw new Error('Failed to serialize state');
        }
        localStorage.setItem('taskManager', serializedState);
    } catch (e) {
        console.error('Failed to save state:', e);
        // Возможно, показать уведомление пользователю
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem('taskManager');
        if (saved) {
            state = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('Unable to load from localStorage');
    }
    render();
}

function getCategoryColor(categoryId) {
    const category = state.categories.find(c => c.id === categoryId);
    return category?.color || '#cccccc';
}

function getCategoryName(categoryId) {
    const category = state.categories.find(c => c.id === categoryId);
    return category?.name || 'Без категории';
}

// Модальные окна
function showNewProjectModal() {
    const modal = document.getElementById('projectModal');
    const form = document.getElementById('projectForm');
    const titleEl = document.getElementById('projectModalTitle');
    
    titleEl.textContent = 'Новый проект';
    form.reset();
    document.getElementById('projectId').value = '';
    modal.style.display = 'block';
}

function showNewTaskModal(projectId) {
    const modal = document.getElementById('taskModal');
    const form = document.getElementById('taskForm');
    const titleEl = document.getElementById('taskModalTitle');
    
    titleEl.textContent = 'Новая задача';
    form.reset();
    document.getElementById('taskId').value = '';
    document.getElementById('taskProjectId').value = projectId;

    updateParentTaskSelect();
    updateCategorySelect();
    updateDependenciesSelect();

    modal.style.display = 'block';
}

function showEditTaskModal(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (!task) return;

    const modal = document.getElementById('taskModal');
    const form = document.getElementById('taskForm');
    const titleEl = document.getElementById('taskModalTitle');
    
    titleEl.textContent = 'Редактировать задачу';
    form.reset();

    document.getElementById('taskId').value = task.id;
    document.getElementById('taskProjectId').value = task.projectId;
    document.getElementById('taskName').value = task.name;
    document.getElementById('taskDescription').value = task.description || '';
    document.getElementById('taskPriority').value = task.priority;
    document.getElementById('taskStatus').value = task.status;
    document.getElementById('taskStartDate').value = task.startDate || '';
    document.getElementById('taskDueDate').value = task.dueDate || '';
    document.getElementById('taskTimeEstimate').value = task.timeEstimate || '';

    updateParentTaskSelect(task.id);
    updateCategorySelect();
    updateDependenciesSelect(task.id);

    document.getElementById('taskCategory').value = task.categoryId || '';
    document.getElementById('taskParent').value = task.parentId || '';

    const checklistContainer = document.querySelector('.checklist-items');
    checklistContainer.innerHTML = '';
    if (task.checklist) {
        task.checklist.forEach(item => {
            addChecklistItem(item.text);
        });
    }

    modal.style.display = 'block';
}

function showNewCategoryModal() {
    const modal = document.getElementById('categoryModal');
    const form = document.getElementById('categoryForm');
    form.reset();
    modal.style.display = 'block';
}

function closeModal(modalId) {
    document.getElementById(modalId).style.display = 'none';
}

// Обработчики форм
function handleProjectSubmit(event) {
    event.preventDefault();
    const form = event.target;
    const id = form.projectId.value || generateId();
    
    const project = {
        id,
        name: form.projectName.value,
        description: form.projectDescription.value,
        createdAt: new Date().toISOString()
    };

    const index = state.projects.findIndex(p => p.id === id);
    if (index > -1) {
        state.projects[index] = project;
    } else {
        state.projects.push(project);
    }

    saveState();
    render();
    closeModal('projectModal');
}

function validateTaskForm(form) {
    const startDate = form.taskStartDate.value;
    const dueDate = form.taskDueDate.value;
    const timeEstimate = parseFloat(form.taskTimeEstimate.value);

    if (startDate && dueDate && new Date(dueDate) < new Date(startDate)) {
        alert('Срок выполнения не может быть раньше даты начала');
        return false;
    }

    if (timeEstimate < 0) {
        alert('Оценка времени не может быть отрицательной');
        return false;
    }

    return true;
}

function selectTask(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        const content = document.getElementById('content');
        content.innerHTML = `
            <div class="task-detail">
                <h2>${task.name}</h2>
                <p>${task.description || 'Нет описания'}</p>
                <div class="task-details">
                    <p><strong>Статус:</strong> ${getStatusLabel(task.status)}</p>
                    <p><strong>Приоритет:</strong> ${getPriorityLabel(task.priority)}</p>
                    ${task.startDate ? `<p><strong>Дата начала:</strong> ${new Date(task.startDate).toLocaleDateString()}</p>` : ''}
                    ${task.dueDate ? `<p><strong>Срок выполнения:</strong> ${new Date(task.dueDate).toLocaleDateString()}</p>` : ''}
                    <button class="btn btn-primary" onclick="showEditTaskModal('${task.id}')">Редактировать</button>
                </div>
                ${renderTaskAttachments(task)}
                ${renderTaskChecklist(task)}
                ${renderTaskDependencies(task)}
                ${renderTaskComments(task)}
            </div>
        `;
    }
}

function handleTaskSubmit(event) {
    event.preventDefault();
    const form = event.target;
    
    if (!validateTaskForm(form)) return;

    const id = form.taskId.value || generateId();
    
    const checklistItems = Array.from(document.querySelectorAll('.checklist-item input[type="text"]'))
        .map(input => ({
            id: generateId(),
            text: input.value,
            completed: false
        }));

    const task = {
        id,
        projectId: form.taskProjectId.value,
        name: form.taskName.value,
        description: form.taskDescription.value,
        categoryId: form.taskCategory.value || null,
        parentId: form.taskParent.value || null,
        priority: form.taskPriority.value,
        status: form.taskStatus.value,
        startDate: form.taskStartDate.value || null,
        dueDate: form.taskDueDate.value || null,
        timeEstimate: parseFloat(form.taskTimeEstimate.value) || 0,
        timeSpent: 0,
        progress: 0,
        checklist: checklistItems,
        dependencies: Array.from(form.taskDependencies.selectedOptions).map(opt => opt.value),
        attachments: [],
        comments: [],
        history: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    };

    const files = form.taskAttachments.files;
    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
            task.attachments.push({
                id: generateId(),
                name: file.name,
                type: file.type,
                url: e.target.result,
                createdAt: new Date().toISOString()
            });
        };
        reader.readAsDataURL(file);
    });

    const index = state.tasks.findIndex(t => t.id === id);
    if (index > -1) {
        state.tasks[index] = task;
    } else {
        state.tasks.push(task);
    }

    saveState();
    render();
    closeModal('taskModal');
}

function handleCategorySubmit(event) {
    event.preventDefault();
    const form = event.target;
    
    const category = {
        id: generateId(),
        name: form.categoryName.value,
        color: form.categoryColor.value,
        projectId: state.currentProject.id
    };

    state.categories.push(category);
    saveState();
    updateCategorySelect();
    closeModal('categoryModal');
}

// Управление задачами
function selectProject(id) {
    state.currentProject = state.projects.find(p => p.id === id) || null;
    render(); // Добавлен принудительный рендер
}

function deleteProject(projectId) {
    if (confirm('Вы уверены, что хотите удалить этот проект? Все задачи будут удалены!')) {
        state.projects = state.projects.filter(p => p.id !== projectId);
        state.tasks = state.tasks.filter(t => t.projectId !== projectId);
        if (state.currentProject?.id === projectId) state.currentProject = null;
        saveState();
        render();
    }
}

function updateTaskStatus(taskId, status) {
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        task.status = status;
        task.updatedAt = new Date().toISOString();
        saveState();
        render();
    }
}

function toggleTaskCompletion(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        task.status = task.status === 'completed' ? 'inProgress' : 'completed';
        task.progress = task.status === 'completed' ? 100 : 0;
        task.updatedAt = new Date().toISOString();
        saveState();
        render();
    }
}

function deleteTask(taskId) {
    if (confirm('Вы уверены, что хотите удалить эту задачу?')) {
        const subtasks = state.tasks.filter(t => t.parentId === taskId);
        subtasks.forEach(st => deleteTask(st.id));
        state.tasks = state.tasks.filter(t => t.id !== taskId);
        saveState();
        render();
    }
}

// Чек-лист
function addChecklistItem(text = '') {
    const container = document.querySelector('.checklist-items');
    const itemId = generateId();
    const itemHtml = `
        <div class="checklist-item" data-id="${itemId}">
            <input type="text" value="${text}" placeholder="Новый пункт" data-id="${itemId}">
            <button type="button" class="btn btn-small" onclick="removeChecklistItem('${itemId}')">×</button>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', itemHtml);
}

function removeChecklistItem(itemId) {
    const item = document.querySelector(`.checklist-item[data-id="${itemId}"]`);
    if (item) item.remove();
}

// Обновление селектов
function updateParentTaskSelect(currentTaskId = null) {
    const select = document.getElementById('taskParent');
    const projectTasks = state.tasks.filter(t => 
        t.projectId === state.currentProject?.id && 
        t.id !== currentTaskId
    );
    select.innerHTML = '<option value="">Нет</option>' + projectTasks.map(task => `
        <option value="${task.id}">${task.name}</option>
    `).join('');
}

function updateCategorySelect() {
    const select = document.getElementById('taskCategory');
    if (!state.categories || !Array.isArray(state.categories)) {
        console.error('Categories is undefined or not an array');
        select.innerHTML = '<option value="">Без категории</option>';
        return;
    }
    const categories = state.categories.filter(c => c.projectId === state.currentProject?.id);
    select.innerHTML = '<option value="">Без категории</option>' + categories.map(category => `
        <option value="${category.id}">${category.name}</option>
    `).join('');
}


function updateDependenciesSelect(currentTaskId = null) {
    const select = document.getElementById('taskDependencies');
    const projectTasks = state.tasks.filter(t => 
        t.projectId === state.currentProject?.id && 
        t.id !== currentTaskId
    );
    select.innerHTML = projectTasks.map(task => `
        <option value="${task.id}">${task.name}</option>
    `).join('');
}

// Визуализация
function render() {
    renderProjectTree();
    renderContent();
}

function renderProjectTree() {
    const tree = document.getElementById('projectTree');
    if (!tree) return;
    
    tree.innerHTML = state.projects.map(project => `
        <div class="tree-item">
            <div class="tree-content ${state.currentProject?.id === project.id ? 'active' : ''}"
                 onclick="selectProject('${project.id}')">
                ${project.name}
                <span class="task-count">
                    ${state.tasks.filter(t => t.projectId === project.id).length}
                </span>
                <div class="project-actions">
                    <button class="btn btn-small" onclick="event.stopPropagation(); deleteProject('${project.id}')">×</button>
                </div>
            </div>
            ${state.currentProject?.id === project.id ?
                `<div class="tree-children">
                    ${renderProjectTasks(project.id)}
                </div>` : ''}
        </div>
    `).join('');
}

function renderProjectTasks(projectId) {
    return state.tasks
        .filter(task => task.projectId === projectId)
        .map(task => `
            <div class="tree-item">
                <div class="tree-content" onclick="selectTask('${task.id}')">
                    ${task.name}
                    ${task.status === 'completed' ? '✓' : ''}
                </div>
            </div>
        `).join('');
}

function renderSubtasks(parentId) {
    return state.tasks
        .filter(task => task.parentId === parentId)
        .map(task => `
            <div class="tree-item">
                <div class="tree-content" onclick="selectTask('${task.id}')">
                    ${task.name}
                    ${task.status === 'completed' ? '✓' : ''}
                </div>
            </div>
        `).join('');
}

function renderTasks(tasks) {
    if (!tasks.length) {
        return '<p>Нет задач</p>';
    }

    return tasks.map(task => `
        <div class="task-item">
            <div style="display: flex; justify-content: space-between">
                <h3>${task.name}</h3>
                <span class="badge badge-${getPriorityClass(task.priority)}">
                    ${getPriorityLabel(task.priority)}
                </span>
            </div>
            <p>${task.description || ''}</p>
            <div class="progress">
                <div class="progress-bar" style="width: ${task.progress}%"></div>
            </div>
            <button class="btn btn-small" onclick="showEditTaskModal('${task.id}')">✎</button>
            <button class="btn btn-small" onclick="deleteTask('${task.id}')">×</button>
        </div>
    `).join('');
}

function showEditTaskModal(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (!task) return;

    const modal = document.getElementById('taskModal');
    const form = document.getElementById('taskForm');
    const titleEl = document.getElementById('taskModalTitle');
    
    titleEl.textContent = 'Редактировать задачу';
    form.reset();

    document.getElementById('taskId').value = task.id;
    document.getElementById('taskProjectId').value = task.projectId;
    document.getElementById('taskName').value = task.name;
    document.getElementById('taskDescription').value = task.description || '';
    document.getElementById('taskPriority').value = task.priority;
    document.getElementById('taskStatus').value = task.status;
    document.getElementById('taskStartDate').value = task.startDate || '';
    document.getElementById('taskDueDate').value = task.dueDate || '';
    document.getElementById('taskTimeEstimate').value = task.timeEstimate || '';

    modal.style.display = 'block';
}

function deleteTask(taskId) {
    if (confirm('Вы уверены, что хотите удалить эту задачу?')) {
        state.tasks = state.tasks.filter(t => t.id !== taskId);
        saveState();
        render();
    }
}

function renderContent() {
    const content = document.getElementById('content');
    if (!content) return;
    
    if (!state.currentProject) {
        content.innerHTML = '<p>Выберите проект или создайте новый</p>';
        return;
    }

    const tasks = state.tasks.filter(t => t.projectId === state.currentProject.id);
    const completedTasks = tasks.filter(t => t.status === 'completed').length;
    const progress = tasks.length ? Math.round((completedTasks / tasks.length) * 100) : 0;

    content.innerHTML = `
        <div class="project-header">
            <div style="display: flex; justify-content: space-between; align-items: start">
                <div>
                    <h2>${state.currentProject.name}</h2>
                    <p>${state.currentProject.description || ''}</p>
                </div>
                <button class="btn btn-primary" onclick="showNewTaskModal('${state.currentProject.id}')">
                    + Добавить задачу
                </button>
            </div>
            <div class="progress" style="margin: 1rem 0">
                <div class="progress-bar" style="width: ${progress}%"></div>
            </div>
            <p style="text-align: right; font-size: 0.875rem">
                Прогресс: ${progress}% (${completedTasks}/${tasks.length})
            </p>
        </div>
        <div class="task-list">
            ${renderTasks(tasks)}
        </div>
    `;
}

// Инициализация приложения
// В единственном оставшемся обработчике DOMContentLoaded
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    
    document.querySelectorAll('.modal').forEach(modal => {
        modal.addEventListener('click', (event) => {
            if (event.target === modal) {
                closeModal(modal.id);
            }
        });
    });

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            document.querySelectorAll('.modal').forEach(modal => {
                if (modal.style.display === 'block') {
                    closeModal(modal.id);
                }
            });
        }
    });
});

// Вспомогательные функции для отображения
function getPriorityClass(priority) {
    switch (priority) {
        case 'high': return 'danger';
        case 'medium': return 'warning';
        case 'low': return 'primary';
        default: return 'primary';
    }
}

function getPriorityLabel(priority) {
    switch (priority) {
        case 'high': return 'Высокий';
        case 'medium': return 'Средний';
        case 'low': return 'Низкий';
        default: return 'Низкий';
    }
}

function getStatusClass(status) {
    switch (status) {
        case 'completed': return 'success';
        case 'inProgress': return 'primary';
        case 'onHold': return 'warning';
        default: return 'secondary';
    }
}

function getStatusLabel(status) {
    switch (status) {
        case 'new': return 'Новая';
        case 'inProgress': return 'В работе';
        case 'onHold': return 'На паузе';
        case 'completed': return 'Завершена';
        default: return 'Новая';
    }
}

function switchView(viewName) {
    state.views.current = viewName;
    renderContent(); // Измените с render() на renderContent()
}

// Рендеринг текущего вида
function renderCurrentView() {
    switch (state.views.current) {
        case 'kanban': return renderKanbanBoard();
        case 'timeline': return renderTimeline();
        case 'stats': return renderStatistics();
        default: return renderTaskTree();
    }
}

function renderTaskTree() {
    const tasks = filterTasks(state.tasks.filter(t => 
        t.projectId === state.currentProject?.id && !t.parentId
    ));
    return tasks.length ? `
        <div class="filter-panel">${renderFilters()}</div>
        <div class="task-list">${tasks.map(task => renderTaskItem(task)).join('')}</div>
    ` : '<p>Нет задач</p>';
}

function renderTaskItem(task) {
    return `
        <div class="task-item">
            <div class="task-header">
                <div class="task-header-content">
                    <h3 class="task-title">${task.name}</h3>
                    <div class="task-badges">
                        <span class="badge badge-${getPriorityClass(task.priority)}">
                            ${getPriorityLabel(task.priority)}
                        </span>
                        <span class="badge badge-${getStatusClass(task.status)}">
                            ${getStatusLabel(task.status)}
                        </span>
                        ${task.categoryId ? `
                            <span class="badge" style="background-color: ${getCategoryColor(task.categoryId)}">
                                ${getCategoryName(task.categoryId)}
                            </span>
                        ` : ''}
                    </div>
                </div>
                <div class="task-actions">
                    <button class="btn btn-small" onclick="toggleTaskCompletion('${task.id}')">
                        ${task.status === 'completed' ? '✓' : '⬜'}
                    </button>
                    <button class="btn btn-small" onclick="showEditTaskModal('${task.id}')">✎</button>
                    <button class="btn btn-small" onclick="deleteTask('${task.id}')">×</button>
                </div>
            </div>
            <div class="task-content">
                ${task.description ? `<p>${task.description}</p>` : ''}
                <div class="task-dates">
                    ${task.startDate ? `<div>Начало: ${new Date(task.startDate).toLocaleDateString()}</div>` : ''}
                    ${task.dueDate ? `<div>Срок: ${new Date(task.dueDate).toLocaleDateString()}</div>` : ''}
                </div>
                ${renderTaskAttachments(task)}
                ${renderTaskChecklist(task)}
                ${renderTaskDependencies(task)}
                ${renderTaskComments(task)}
            </div>
            <div class="task-children">
                ${renderSubtasks(task.id)}
            </div>
        </div>
    `;
}

function renderTaskAttachments(task) {
    if (!task.attachments?.length) return '';
    return `
        <div class="task-attachments">
            <h4>Вложения</h4>
            <div class="attachments-list">
                ${task.attachments.map(file => `
                    <div class="attachment-item">
                        <a href="${file.url}" target="_blank" download="${file.name}">
                            ${file.name}
                        </a>
                        <span class="attachment-date">
                            ${new Date(file.createdAt).toLocaleDateString()}
                        </span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

function renderTaskChecklist(task) {
    if (!task.checklist?.length) return '';
    return `
        <div class="task-checklist">
            <h4>Чек-лист</h4>
            <div class="checklist-items">
                ${task.checklist.map(item => `
                    <div class="checklist-item">
                        <input type="checkbox" ${item.completed ? 'checked' : ''} 
                               onchange="toggleChecklistItem('${task.id}', '${item.id}')">
                        <span>${item.text}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

function renderTaskDependencies(task) {
    if (!task.dependencies?.length) return '';
    return `
        <div class="task-dependencies">
            <h4>Зависимости</h4>
            <div class="dependencies-list">
                ${task.dependencies.map(depId => {
                    const depTask = state.tasks.find(t => t.id === depId);
                    return depTask ? `
                        <div class="dependency-item">
                            <span>${depTask.name}</span>
                            <span class="badge badge-${getStatusClass(depTask.status)}">
                                ${getStatusLabel(depTask.status)}
                            </span>
                        </div>
                    ` : '';
                }).join('')}
            </div>
        </div>
    `;
}

function renderTaskComments(task) {
    if (!task.comments?.length) return '';
    return `
        <div class="task-comments">
            <h4>Комментарии</h4>
            <div class="comments-list">
                ${task.comments.map(comment => `
                    <div class="comment-item">
                        <div class="comment-header">
                            <span class="comment-date">
                                ${new Date(comment.createdAt).toLocaleString()}
                            </span>
                        </div>
                        <div class="comment-content">
                            ${comment.text}
                        </div>
                    </div>
                `).join('')}
            </div>
            <div class="comment-form">
                <textarea placeholder="Добавить комментарий..." 
                          onkeydown="if(event.key === 'Enter' && !event.shiftKey) { 
                              event.preventDefault(); 
                              addComment('${task.id}', this.value); 
                              this.value = ''; 
                          }"></textarea>
            </div>
        </div>
    `;
}

function toggleChecklistItem(taskId, itemId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        const item = task.checklist.find(i => i.id === itemId);
        if (item) {
            item.completed = !item.completed;
            saveState();
            render();
        }
    }
}

function addComment(taskId, text) {
    if (!text.trim()) return;
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        task.comments = task.comments || [];
        task.comments.push({
            id: generateId(),
            text: text.trim(),
            createdAt: new Date().toISOString()
        });
        saveState();
        render();
    }
}

// Рендеринг канбан-доски
function renderKanbanBoard() {
    const statuses = ['new', 'inProgress', 'onHold', 'completed'];
    const tasks = filterTasks(state.tasks.filter(t => t.projectId === state.currentProject?.id));

    return `
        <div class="filter-panel">
            ${renderFilters()}
        </div>
        <div class="kanban-board">
            ${statuses.map(status => `
                <div class="kanban-column">
                    <div class="kanban-column-header">
                        <h3>${getStatusLabel(status)}</h3>
                        <span class="task-count">
                            ${tasks.filter(t => t.status === status).length}
                        </span>
                    </div>
                    <div class="kanban-column-content" 
                         ondrop="dropTask(event)" 
                         ondragover="allowDrop(event)" 
                         data-status="${status}">
                        ${tasks.filter(t => t.status === status)
                            .map(task => renderKanbanCard(task))
                            .join('')}
                    </div>
                </div>
            `).join('')}
        </div>
    `;
}

function renderKanbanCard(task) {
    return `
        <div class="kanban-card" 
             draggable="true" 
             ondragstart="dragTask(event)" 
             data-task-id="${task.id}">
            <div class="kanban-card-header">
                <h4>${task.name}</h4>
                <div class="task-badges">
                    <span class="badge badge-${getPriorityClass(task.priority)}">
                        ${getPriorityLabel(task.priority)}
                    </span>
                    ${task.categoryId ? `
                        <span class="badge" style="background-color: ${getCategoryColor(task.categoryId)}">
                            ${getCategoryName(task.categoryId)}
                        </span>
                    ` : ''}
                </div>
            </div>
            ${task.dueDate ? `
                <div class="kanban-card-date">
                    Срок: ${new Date(task.dueDate).toLocaleDateString()}
                </div>
            ` : ''}
            <div class="progress">
                <div class="progress-bar" style="width: ${getTaskProgress(task)}%"></div>
            </div>
        </div>
    `;
}

// Функции для Drag & Drop
function dragTask(event) {
    event.dataTransfer.setData('taskId', event.target.dataset.taskId);
}

function allowDrop(event) {
    event.preventDefault();
}

function dropTask(event) {
    event.preventDefault();
    const taskId = event.dataTransfer.getData('taskId');
    const newStatus = event.target.closest('.kanban-column-content').dataset.status;
    
    const task = state.tasks.find(t => t.id === taskId);
    if (task && task.status !== newStatus) {
        const oldStatus = task.status;
        task.status = newStatus;
        task.updatedAt = new Date().toISOString();

        task.history = task.history || [];
        task.history.push({
            id: generateId(),
            type: 'status_change',
            oldValue: oldStatus,
            newValue: newStatus,
            timestamp: new Date().toISOString()
        });

        saveState();
        render();
    }
}

// Компоненты фильтрации
function renderFilters() {
    return `
        <div class="filters">
            <div class="filter-group">
                <label>Категория</label>
                <select onchange="updateFilter('category', this.value)">
                    <option value="">Все категории</option>
                    ${state.categories
                        .filter(c => c.projectId === state.currentProject?.id)
                        .map(category => `
                            <option value="${category.id}" 
                                    ${state.views.filters.category === category.id ? 'selected' : ''}>
                                ${category.name}
                            </option>
                        `).join('')}
                </select>
            </div>
            <div class="filter-group">
                <label>Приоритет</label>
                <select onchange="updateFilter('priority', this.value)">
                    <option value="">Все приоритеты</option>
                    <option value="high" ${state.views.filters.priority === 'high' ? 'selected' : ''}>
                        Высокий
                    </option>
                    <option value="medium" ${state.views.filters.priority === 'medium' ? 'selected' : ''}>
                        Средний
                    </option>
                    <option value="low" ${state.views.filters.priority === 'low' ? 'selected' : ''}>
                        Низкий
                    </option>
                </select>
            </div>
            <div class="filter-group">
                <label>Статус</label>
                <select onchange="updateFilter('status', this.value)">
                    <option value="">Все статусы</option>
                    <option value="new" ${state.views.filters.status === 'new' ? 'selected' : ''}>
                        Новые
                    </option>
                    <option value="inProgress" ${state.views.filters.status === 'inProgress' ? 'selected' : ''}>
                        В работе
                    </option>
                    <option value="onHold" ${state.views.filters.status === 'onHold' ? 'selected' : ''}>
                        На паузе
                    </option>
                    <option value="completed" ${state.views.filters.status === 'completed' ? 'selected' : ''}>
                        Завершенные
                    </option>
                </select>
            </div>
            <div class="filter-group">
                <label>Поиск</label>
                <input type="text" 
                       placeholder="Поиск по названию..." 
                       value="${state.views.filters.search}"
                       oninput="updateFilter('search', this.value)">
            </div>
        </div>
    `;
}

function updateFilter(type, value) {
    state.views.filters[type] = value;
    render();
}

function filterTasks(tasks) {
    const filters = state.views.filters;
    return tasks.filter(task => {
        if (filters.category && task.categoryId !== filters.category) return false;
        if (filters.priority && task.priority !== filters.priority) return false;
        if (filters.status && task.status !== filters.status) return false;
        if (filters.search && !task.name.toLowerCase().includes(filters.search.toLowerCase())) return false;
        return true;
    });
}

// Рендеринг временной шкалы
function renderTimeline() {
    const tasks = filterTasks(state.tasks.filter(t => 
        t.projectId === state.currentProject?.id &&
        (t.startDate || t.dueDate)
    ));

    if (!tasks.length) {
        return '<p>Нет задач с установленными сроками</p>';
    }

    const dates = tasks.flatMap(t => [
        t.startDate && new Date(t.startDate),
        t.dueDate && new Date(t.dueDate)
    ].filter(Boolean));

    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    minDate.setDate(minDate.getDate() - 2);
    maxDate.setDate(maxDate.getDate() + 2);

    return `
        <div class="filter-panel">
            ${renderFilters()}
        </div>
        <div class="timeline">
            <div class="timeline-header">
                ${renderTimelineHeader(minDate, maxDate)}
            </div>
            <div class="timeline-body">
                ${renderTimelineTasks(tasks, minDate, maxDate)}
            </div>
        </div>
    `;
}

function renderTimelineHeader(minDate, maxDate) {
    const days = [];
    const currentDate = new Date(minDate);
    
    while (currentDate <= maxDate) {
        days.push(new Date(currentDate));
        currentDate.setDate(currentDate.getDate() + 1);
    }

    return `
        <div class="timeline-scale">
            ${days.map(date => `
                <div class="timeline-day">
                    ${date.toLocaleDateString()}
                </div>
            `).join('')}
        </div>
    `;
}

function renderTimelineTasks(tasks, minDate, maxDate) {
    const totalDays = (maxDate - minDate) / (1000 * 60 * 60 * 24);

    return tasks.map(task => {
        const start = task.startDate ? new Date(task.startDate) : minDate;
        const end = task.dueDate ? new Date(task.dueDate) : maxDate;
        
        const startOffset = (start - minDate) / (1000 * 60 * 60 * 24);
        const duration = (end - start) / (1000 * 60 * 60 * 24);
        
        const leftPercent = (startOffset / totalDays) * 100;
        const widthPercent = (duration / totalDays) * 100;

        return `
            <div class="timeline-task">
                <div class="timeline-task-info" style="width: 200px;">
                    <div class="task-title">${task.name}</div>
                    <div class="task-badges">
                        <span class="badge badge-${getPriorityClass(task.priority)}">
                            ${getPriorityLabel(task.priority)}
                        </span>
                    </div>
                </div>
                <div class="timeline-task-bar" 
                     style="left: calc(200px + ${leftPercent}%); width: ${widthPercent}%"
                     title="${task.name}">
                    <div class="progress-bar" style="width: ${getTaskProgress(task)}%"></div>
                </div>
            </div>
        `;
    }).join('');
}

// Рендеринг статистики
function renderStatistics() {
    const tasks = state.tasks.filter(t => t.projectId === state.currentProject?.id);
    
    const totalTasks = tasks.length;
    const completedTasks = tasks.filter(t => t.status === 'completed').length;
    const totalEstimated = tasks.reduce((sum, t) => sum + (t.timeEstimate || 0), 0);
    const totalSpent = tasks.reduce((sum, t) => sum + (t.timeSpent || 0), 0);
    
    const overdueTasks = tasks.filter(t => 
        t.dueDate && 
        new Date(t.dueDate) < new Date() && 
        t.status !== 'completed'
    ).length;

    const byPriority = tasks.reduce((acc, task) => {
        acc[task.priority] = (acc[task.priority] || 0) + 1;
        return acc;
    }, {});

    const byStatus = tasks.reduce((acc, task) => {
        acc[task.status] = (acc[task.status] || 0) + 1;
        return acc;
    }, {});

    const byCategory = tasks.reduce((acc, task) => {
        if (task.categoryId) {
            acc[task.categoryId] = (acc[task.categoryId] || 0) + 1;
        }
        return acc;
    }, {});

    return `
        <div class="statistics">
            <div class="stats-grid">
                <div class="stats-card">
                    <h3>Общий прогресс</h3>
                    <div class="stats-number">
                        ${totalTasks ? Math.round((completedTasks / totalTasks) * 100) : 0}%
                    </div>
                    <div class="progress" style="height: 20px;">
                        <div class="progress-bar" 
                             style="width: ${totalTasks ? (completedTasks/totalTasks*100) : 0}%">
                        </div>
                    </div>
                    <p>${completedTasks} из ${totalTasks} задач завершено</p>
                </div>

                <div class="stats-card">
                    <h3>Просроченные задачи</h3>
                    <div class="stats-number ${overdueTasks ? 'text-danger' : ''}">
                        ${overdueTasks}
                    </div>
                    <p>задач требуют внимания</p>
                </div>

                <div class="stats-card">
                    <h3>Затраченное время</h3>
                    <div class="stats-number">
                        ${totalSpent}ч / ${totalEstimated}ч
                    </div>
                    <div class="progress" style="height: 20px;">
                        <div class="progress-bar" 
                             style="width: ${totalEstimated ? (totalSpent/totalEstimated*100) : 0}%">
                        </div>
                    </div>
                </div>
            </div>

            <div class="stats-details">
                <div class="stats-section">
                    <h3>По приоритетам</h3>
                    <div class="stats-bars">
                        ${Object.entries(byPriority).map(([priority, count]) => `
                            <div class="stats-bar">
                                <div class="stats-bar-label">
                                    ${getPriorityLabel(priority)}
                                </div>
                                <div class="stats-bar-value badge-${getPriorityClass(priority)}">
                                    ${count}
                                </div>
                            </div>
                        `).join('')}
                    </div>
                </div>

                <div class="stats-section">
                    <h3>По статусам</h3>
                    <div class="stats-bars">
                        ${Object.entries(byStatus).map(([status, count]) => `
                            <div class="stats-bar">
                                <div class="stats-bar-label">
                                    ${getStatusLabel(status)}
                                </div>
                                <div class="stats-bar-value badge-${getStatusClass(status)}">
                                    ${count}
                                </div>
                            </div>
                        `).join('')}
                    </div>
                </div>

                ${Object.keys(byCategory).length ? `
                    <div class="stats-section">
                        <h3>По категориям</h3>
                        <div class="stats-bars">
                            ${Object.entries(byCategory).map(([categoryId, count]) => `
                                <div class="stats-bar">
                                    <div class="stats-bar-label">
                                        ${getCategoryName(categoryId)}
                                    </div>
                                    <div class="stats-bar-value" 
                                         style="background-color: ${getCategoryColor(categoryId)}">
                                        ${count}
                                    </div>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                ` : ''}
            </div>

            <div class="stats-section">
                <h3>Распределение времени</h3>
                <div class="time-distribution">
                    ${renderTimeDistribution(tasks)}
                </div>
            </div>
        </div>
    `;
}

// Вспомогательная функция для расчета распределения времени
function renderTimeDistribution(tasks) {
    const timeData = tasks.reduce((acc, task) => {
        if (task.timeSpent) {
            const date = task.startDate ? new Date(task.startDate).toLocaleDateString() : 'Без даты';
            acc[date] = (acc[date] || 0) + task.timeSpent;
        }
        return acc;
    }, {});

    if (Object.keys(timeData).length === 0) {
        return '<p>Нет данных о затраченном времени</p>';
    }

    return `
        <div class="time-chart">
            ${Object.entries(timeData).map(([date, hours]) => `
                <div class="time-bar">
                    <div class="time-bar-label">${date}</div>
                    <div class="time-bar-value" style="width: ${Math.min(hours * 5, 100)}%">
                        ${hours}ч
                    </div>
                </div>
            `).join('')}
        </div>
    `;
}

// Функция для расчета прогресса задачи
function calculateTaskProgress(task) {
    if (task.status === 'completed') return 100;
    
    let progress = 0;
    let factors = 0;

    // Учитываем чек-лист
    if (task.checklist?.length) {
        const completed = task.checklist.filter(item => item.completed).length;
        progress += (completed / task.checklist.length) * 100;
        factors++;
    }

    // Учитываем затраченное время
    if (task.timeEstimate && task.timeSpent) {
        const timeProgress = Math.min((task.timeSpent / task.timeEstimate) * 100, 100);
        progress += timeProgress;
        factors++;
    }

    // Учитываем дату выполнения
    if (task.startDate && task.dueDate) {
        const total = new Date(task.dueDate) - new Date(task.startDate);
        const elapsed = new Date() - new Date(task.startDate);
        const dateProgress = Math.min((elapsed / total) * 100, 100);
        progress += dateProgress;
        factors++;
    }

    return factors ? Math.round(progress / factors) : 0;
}

// Функция для получения прогресса задачи
function getTaskProgress(task) {
    return calculateTaskProgress(task);
}
    
// Обработчики событий для модальных окон
document.querySelectorAll('.modal').forEach(modal => {
        modal.addEventListener('click', (event) => {
            if (event.target === modal) {
                closeModal(modal.id);
            }
        });
});

// Обработка клавиши Escape
document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            document.querySelectorAll('.modal').forEach(modal => {
                if (modal.style.display === 'block') {
                    closeModal(modal.id);
                }
            });
        }
});
