# Conversation Archive Browser

A web-based tool for browsing, searching, and analyzing exported conversation archives with an intuitive interface.

## Overview

This application provides a comprehensive solution for managing and exploring conversation archives. It features advanced search capabilities, automated content categorization, and export functionality.

## Features

### Core Functionality
- **Archive Import**: Support for ZIP and JSON file formats
- **Conversation Parsing**: Automatic extraction and organization of conversation threads
- **Advanced Search**: Full-text search across titles and message content
- **Smart Filtering**: Multiple filter criteria can be combined for precise results

### Content Analysis
- **Automated Categorization**: Intelligent content classification across 12 distinct categories
- **Entity Detection**: Automatic identification and tracking of named entities
- **Analytics Dashboard**: Statistical overview of conversation participants and themes

### Export Capabilities
- **Document Generation**: Export conversations to Word (.docx) format
- **PDF Creation**: Generate formatted PDF documents
- **Formatted Output**: Color-coded messages with timestamps and metadata

### User Interface
- **Responsive Design**: Modern, clean interface with custom styling
- **Sidebar Navigation**: Easy access to analytics and filtering options
- **Expandable Views**: Detailed conversation viewing on demand
- **Sort Options**: Multiple sorting criteria (date, title, message count)

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Setup

1. Clone the repository:
```bash
git clone <repository-url>
cd searchchatfunction
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`

## Usage

### Importing Archives

1. Navigate to the "Upload & Browse" section
2. Click the upload button and select your archive file (ZIP or JSON)
3. Wait for processing to complete
4. Review the success message showing the number of imported conversations

### Searching and Filtering

- **Text Search**: Enter keywords to search across conversation titles or content
- **Category Filter**: Select one or more categories to filter conversations
- **Entity Search**: Filter by specific named entities (comma-separated)
- **Combined Filters**: All filters work together for refined results

### Viewing Conversations

1. Click the "View" button on any conversation card
2. Review the full message thread
3. Export to Word or PDF as needed

### Analytics

The sidebar "Analytics" section provides:
- Entity frequency analysis
- Category distribution
- Thread participation metrics
- Quick filtering by entity

## File Format Support

The application supports ChatGPT export formats:
- Standard JSON export files
- ZIP archives containing conversation data

## Privacy and Security

- All processing occurs locally within the application
- No data is transmitted to external servers
- Uploaded files are processed in-memory only

## Technical Details

### Built With
- **Streamlit**: Web application framework
- **python-docx**: Word document generation
- **ReportLab**: PDF creation
- **Python Standard Library**: JSON parsing, file handling

### Architecture
- Client-side processing
- In-memory data management
- Stateful session handling

## Configuration

The application includes several pre-configured settings:
- 12 content categories with keyword-based detection
- Entity recognition patterns
- Export formatting templates
- UI theme and styling

## Troubleshooting

### Common Issues

**Upload fails**: Ensure the file is a valid JSON or ZIP archive
**No conversations found**: Verify the file structure matches ChatGPT export format
**Performance issues**: Large archives (1000+ conversations) may take time to process

### Performance Tips

- Use specific search terms to reduce result sets
- Apply filters before viewing conversations
- Export individual conversations rather than bulk exports

## Contributing

This is a personal project. If you encounter issues, please report them through the repository's issue tracker.

## License

This project is provided as-is for personal use.

## Acknowledgments

Built with Streamlit and Python standard libraries.
