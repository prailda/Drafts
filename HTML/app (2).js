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

// Вспомогательные функции
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2);
}

function saveState() {
    try {
        localStorage.setItem('taskManager', JSON.stringify(state));
    } catch (e) {
        console.warn('Unable to save to localStorage');
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

    // Заполняем поля формы
    document.getElementById('taskId').value = task.id;
    document.getElementById('taskProjectId').value = task.projectId;
    document.getElementById('taskName').value = task.name;
    document.getElementById('taskDescription').value = task.description || '';
    document.getElementById('taskPriority').value = task.priority;
    document.getElementById('taskStatus').value = task.status;
    document.getElementById('taskStartDate').value = task.startDate || '';
    document.getElementById('taskDueDate').value = task.dueDate || '';
    document.getElementById('taskTimeEstimate').value = task.timeEstimate || '';

    // Заполняем списки выбора
    updateParentTaskSelect(task.id);
    updateCategorySelect();
    updateDependenciesSelect(task.id);

    // Устанавливаем значения
    document.getElementById('taskCategory').value = task.categoryId || '';
    document.getElementById('taskParent').value = task.parentId || '';

    // Заполняем чек-лист
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

function handleTaskSubmit(event) {
    event.preventDefault();
    const form = event.target;
    const id = form.taskId.value || generateId();
    
    // Собираем чек-лист
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

    // Обработка файлов
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
    state.currentProject = state.projects.find(p => p.id === id);
    render();
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
        // Удаляем все подзадачи
        const subtasks = state.tasks.filter(t => t.parentId === taskId);
        subtasks.forEach(st => deleteTask(st.id));

        // Удаляем задачу
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

    select.innerHTML = '<option value="">Нет</option>' +
        projectTasks.map(task => `
            <option value="${task.id}">${task.name}</option>
        `).join('');
}

function updateCategorySelect() {
    const select = document.getElementById('taskCategory');
    const categories = state.categories.filter(c => 
        c.projectId === state.currentProject?.id
    );

    select.innerHTML = '<option value="">Без категории</option>' +
        categories.map(category => `
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
            </div>
            <div class="tree-children">
                ${renderProjectTasks(project.id)}
            </div>
        </div>
    `).join('');
}

function renderProjectTasks(projectId) {
    const tasks = state.tasks
        .filter(task => task.projectId === projectId && !task.parentId)
        .map(task => `
            <div class="tree-item">
                <div class="tree-content">
                    ${task.name}
                    ${task.status === 'completed' ? '✓' : ''}
                </div>
                <div class="tree-children">
                    ${renderSubtasks(task.id)}
                </div>
            </div>
        `).join('');
    return tasks;
}

function renderSubtasks(parentId) {
    return state.tasks
        .filter(task => task.parentId === parentId)
        .map(task => `
            <div class="tree-item">
                <div class="tree-content">
                    ${task.name}
                    ${task.status === 'completed' ? '✓' : ''}
                </div>
            </div>
        `).join('');
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

        <div class="view-selector">
            <button class="btn ${state.views.current === 'tree' ? 'btn-primary' : ''}" 
                    onclick="switchView('tree')">Список</button>
            <button class="btn ${state.views.current === 'kanban' ? 'btn-primary' : ''}" 
                    onclick="switchView('kanban')">Канбан</button>
            <button class="btn ${state.views.current === 'timeline' ? 'btn-primary' : ''}" 
                    onclick="switchView('timeline')">Timeline</button>
            <button class="btn ${state.views.current === 'stats' ? 'btn-primary' : ''}" 
                    onclick="switchView('stats')">Статистика</button>
        </div>

        <div class="view-content">
            ${renderCurrentView()}
        </div>
    `;
}

// Инициализация приложения
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    
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