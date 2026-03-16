# SpreadsheetApp

A modern web application built with React and Vite for creating and managing spreadsheets.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- **Node.js** (version 18 or higher) - [Download](https://nodejs.org/)
- **npm** (comes with Node.js) or **yarn**
- **Git** - [Download](https://git-scm.com/)

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/Nandanmarepally/task_AI_native_Office_intern.git
cd task_AI_native_Office_intern
```

### 2. Install Dependencies

Install all required project dependencies:

```bash
npm install
```

Or if you prefer yarn:

```bash
yarn install
```

### 3. Run Development Server

Start the development server with hot module replacement (HMR):

```bash
npm run dev
```

The application will be available at `http://localhost:5173`

### 4. Build for Production

Create an optimized production build:

```bash
npm run build
```

The build output will be in the `dist/` directory.

### 5. Preview Production Build

Preview the production build locally:

```bash
npm run preview
```

### 6. Lint Code

Run ESLint to check for code quality issues:

```bash
npm run lint
```

## Project Structure

```
SpreadsheetApp/
├── src/
│   ├── App.jsx           # Main React component
│   ├── App.css           # Application styles
│   ├── main.jsx          # Application entry point
│   ├── index.css         # Global styles
│   ├── assets/           # Static assets (images, icons, etc.)
│   └── engine/           # Core application logic
│       └── core.js       # Engine core functionality
├── public/               # Static files served as-is
├── package.json          # Project dependencies and scripts
├── vite.config.js        # Vite configuration
├── eslint.config.js      # ESLint configuration
├── index.html            # HTML entry point
└── README.md             # This file
```

## Technologies Used

- **React** - A JavaScript library for building user interfaces
- **Vite** - A next-generation frontend build tool
- **ESLint** - JavaScript linting utility
- **CSS** - Styling and layout

## Features Implemented

**Column Sorting**

- 3-state sorting: Ascending → Descending → None
- Sorts based on computed values (formula results)
- View-layer filtering only (original data intact)

**Excel-like Filtering**

- Interactive filter dropdowns in column headers
- Multi-column filtering support
- Rows are hidden, not deleted (preserves original data)

**Copy & Paste**

- Copy cells with `Ctrl+C` (copies computed values)
- Paste multi-row/multi-column data with `Ctrl+V`
- Supports pasting from Excel and Google Sheets
- Paste operations are fully undoable with `Ctrl+Z`

**Local Storage Persistence**

- Automatic save every 500ms (debounced)
- Full spreadsheet restoration on page refresh
- Persists cell values, formulas, styles, and grid dimensions
- Graceful error handling for storage quota exceeded
- Auto-recovery from corrupted data

## Video Walkthrough

Watch the 3–5 minute walkthrough explaining the implementation and features:

🔗 **Loom Video:** [https://www.loom.com/share/9642f3a7eedb42639930f40a3b34b3f5](https://www.loom.com/share/9642f3a7eedb42639930f40a3b34b3f5)

## Available Scripts

| Command           | Description                                  |
| ----------------- | -------------------------------------------- |
| `npm run dev`     | Start the development server with hot reload |
| `npm run build`   | Build the application for production         |
| `npm run preview` | Preview the production build locally         |
| `npm run lint`    | Run ESLint to check code quality             |

## Development Workflow

1. Create a new branch for your feature:

   ```bash
   git checkout -b feature/your-feature-name
   ```

2. Make your changes and ensure the code passes linting:

   ```bash
   npm run lint
   ```

3. Commit your changes:

   ```bash
   git commit -m "Add description of your changes"
   ```

4. Push to your fork and create a Pull Request

## Browser Support

This application works on all modern browsers that support ES2020+ JavaScript:

- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)

## Troubleshooting

### Dependencies won't install

- Clear npm cache: `npm cache clean --force`
- Delete `node_modules` and `package-lock.json`, then reinstall: `rm -rf node_modules package-lock.json && npm install`

### Port 5173 already in use

- The dev server will automatically try the next available port
- Or specify a custom port: `npm run dev -- --port 3000`

### Build fails

- Ensure all dependencies are installed: `npm install`
- Clear any build cache: `rm -rf dist`
- Try rebuilding: `npm run build`

# task_AI_native_Office_intern
