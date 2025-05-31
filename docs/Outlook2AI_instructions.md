# Outlook2AI - MS Outlook Email Extraction for LLM Analysis

A Python application that extracts emails from MS Outlook desktop application and creates structured DataFrames optimized for Large Language Model (LLM) analysis.

## Features

- **Direct Outlook Integration**: Connects to MS Outlook desktop application using COM interface
- **Comprehensive Email Extraction**: Extracts all email metadata including body, attachments, recipients
- **LLM-Optimized DataFrames**: Creates structured data perfect for AI analysis
- **Flexible Folder Selection**: Extract from specific folders or entire mailbox
- **Multiple Export Formats**: Support for CSV, JSON, and Parquet formats
- **Rich Metadata**: Includes computed fields for enhanced analysis
- **Windows Compatible**: Designed specifically for Windows laptops with Outlook

## System Requirements

- **Operating System**: Windows 10/11
- **Python**: 3.10 or higher
- **MS Outlook**: Desktop application (part of Microsoft Office)
- **Dependencies**: See requirements.txt

## Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd outlook2ai
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Install the package**:
   ```bash
   pip install -e .
   ```

## Quick Start

### Command Line Usage

1. **List available folders**:
   ```bash
   outlook2ai --list-folders
   ```

2. **Extract emails from Inbox**:
   ```bash
   outlook2ai --folders Inbox --output emails.csv
   ```

3. **Extract from multiple folders with limit**:
   ```bash
   outlook2ai --folders Inbox "Sent Items" --max-emails 100 --format json
   ```

### Python API Usage

```python
from outlook2ai import Outlook2AI

# Initialize the application
app = Outlook2AI()

# Connect to Outlook
if app.connect_to_outlook():
    # List available folders
    folders = app.list_folders()
    print("Available folders:")
    for folder in folders[:5]:  # Show first 5
        print(f"  {folder['path']} ({folder['item_count']} items)")
    
    # Extract emails from specific folders
    folder_paths = ["Inbox", "Sent Items"]
    if app.extract_emails(folder_paths, max_emails_per_folder=50):
        # Get the DataFrame
        df = app.get_dataframe()
        print(f"Extracted {len(df)} emails")
        
        # Show summary statistics
        stats = app.get_summary_statistics()
        print(f"Date range: {stats['date_range']}")
        print(f"Unique senders: {len(stats['sender_distribution'])}")
        
        # Export for analysis
        app.export_data("./data/my_emails.csv", format_type="csv")
        
        # Prepare data for LLM analysis
        llm_data = app.prepare_for_llm(max_emails=20)
        print("Data ready for LLM analysis:")
        print(llm_data[:500] + "..." if len(llm_data) > 500 else llm_data)
    
    # Disconnect when done
    app.disconnect()
```

## DataFrame Schema

The extracted email data includes the following columns:

### Core Email Fields
- `folder_name`: Source folder name
- `subject`: Email subject line
- `sender_email`: Sender's email address
- `sender_name`: Sender's display name
- `received_time`: When email was received
- `sent_time`: When email was sent
- `body_text`: Plain text body content
- `body_html`: HTML body content

### Metadata Fields
- `importance`: Email importance level (1=Low, 2=Normal, 3=High)
- `size`: Email size in bytes
- `unread`: Boolean indicating if email is unread
- `has_attachments`: Boolean indicating presence of attachments
- `attachment_count`: Number of attachments
- `to_recipients`: To recipients (semicolon separated)
- `cc_recipients`: CC recipients (semicolon separated)
- `categories`: Outlook categories

### LLM Analysis Fields
- `body_word_count`: Number of words in body
- `subject_length`: Length of subject line
- `is_reply`: Boolean indicating if email is a reply
- `is_forward`: Boolean indicating if email is forwarded
- `domain`: Sender's email domain
- `hour_received`: Hour when email was received (0-23)
- `day_of_week`: Day of week when received
- `time_category`: Time category (Morning/Afternoon/Evening/Night)
- `age_days`: Age of email in days

## Configuration

Create a `config/config.yaml` file to customize behavior:

```yaml
outlook:
  timeout: 30
  default_folders: ["Inbox"]
  max_emails_per_folder: null

dataframe:
  export_format: "csv"
  include_html_body: false
  clean_text: true

llm:
  max_emails_for_prompt: 100
  include_body_text: true
  max_body_length: 1000

logging:
  level: "INFO"
  file: "logs/outlook2ai.log"
```

## Use Cases for LLM Analysis

The extracted DataFrame is optimized for various LLM analysis tasks:

1. **Email Classification**: Categorize emails by topic, urgency, or type
2. **Sentiment Analysis**: Analyze emotional tone of communications
3. **Entity Extraction**: Identify people, organizations, dates, etc.
4. **Summary Generation**: Create concise summaries of email content
5. **Action Item Detection**: Identify emails requiring follow-up
6. **Communication Pattern Analysis**: Understand email usage patterns
7. **Spam/Phishing Detection**: Identify suspicious emails
8. **Content Analysis**: Extract key themes and topics

## Example LLM Prompts

After exporting your data, you can use it with LLMs like this:

```
Analyze the following email dataset and provide insights:
1. Categorize emails by topic (work, personal, marketing, etc.)
2. Identify emails requiring immediate attention
3. Extract key action items and deadlines
4. Summarize communication patterns by sender domain

[Paste the prepared LLM data here]
```

## Security and Privacy

- **Local Processing**: All email processing happens locally on your machine
- **No Cloud Storage**: Email data never leaves your computer
- **Outlook Integration**: Uses official Microsoft COM interface
- **Read-Only Access**: Application only reads emails, never modifies them
- **Secure Exports**: Exported data can be encrypted or stored securely

## Troubleshooting

### Common Issues

1. **"Failed to connect to Outlook"**
   - Ensure Outlook desktop application is installed and running
   - Run Python script as Administrator if needed
   - Check that Outlook is not in Safe Mode

2. **"Access Denied" errors**
   - Outlook may prompt for security permissions
   - Allow the application to access Outlook data
   - Check Outlook Trust Center settings

3. **Memory issues with large datasets**
   - Use `max_emails_per_folder` parameter to limit extraction
   - Process folders one at a time for very large mailboxes
   - Consider exporting to Parquet format for better compression

4. **COM interface errors**
   - Restart Outlook and try again
   - Ensure no other applications are heavily using Outlook
   - Check Windows Event Logs for detailed error information

### Performance Tips

- **Limit email count** for initial testing
- **Use specific folders** rather than extracting entire mailbox
- **Export to Parquet** for large datasets (better compression and speed)
- **Close other Outlook add-ins** during extraction
- **Use SSD storage** for better I/O performance

## Development

### Running Tests

```bash
pytest tests/ -v --cov=outlook2ai
```

### Code Formatting

```bash
black src/ tests/
flake8 src/ tests/
```

### Building Distribution

```bash
python -m build
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and changes.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Search existing GitHub issues
3. Create a new issue with detailed information about your problem

---

**Note**: This application requires MS Outlook desktop application and is designed for Windows environments. It uses the official Microsoft COM interface for secure, read-only access to email data.