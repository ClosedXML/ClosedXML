# ClosedXML WebApp Example

<img width="1911" height="909" alt="image" src="https://github.com/user-attachments/assets/1404dda9-a66c-4f80-8f40-85fa0e9a2bbe" />

<img width="1916" height="914" alt="image" src="https://github.com/user-attachments/assets/cebe4ce4-b28a-4c1a-8f7d-5eeccc89a6b8" />

Web application demonstrating Excel workbook management using [ClosedXML](https://github.com/ClosedXML/ClosedXML) and [Ivy Framework](https://github.com/Ivy-Interactive/Ivy-Framework).

## Features

**Workbooks Viewer:**

- Browse and preview workbook data with dropdown selection
- Displays first 50 rows for performance

**Workbooks Editor:**

- Create, edit, and delete Excel workbooks
- Add columns (string, int, double, decimal, long) and rows
- Blade-based multi-panel interface
- Auto-save to in-memory storage
- Sample data: Employees, Products, and Sales workbooks

## How to Run

**Prerequisites:** .NET 9.0 SDK

```bash
cd ClosedXML.Examples/WebApp
dotnet restore
dotnet watch
```

Open your browser to the URL shown in the terminal (typically `http://localhost:5010`).

## Docker

Build and run with Docker:

```bash
cd ClosedXML.Examples/WebApp
docker build -t closedxml-webapp .
docker run -p 80:80 closedxml-webapp
```

The Dockerfile uses a multi-stage build with .NET 9.0 runtime and exposes port 80.

## How to Deploy

```bash
cd ClosedXML.Examples/WebApp
ivy deploy
```

## Learn More

- [ClosedXML GitHub](https://github.com/ClosedXML/ClosedXML)
- [Ivy Documentation](https://docs.ivy.app)
