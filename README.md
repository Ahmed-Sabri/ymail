[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue?logo=python)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow)](LICENSE)
[![Status: Stable](https://img.shields.io/badge/status-stable-green)](https://github.com/Ahmed-Sabri/ymail)
[![Issues](https://img.shields.io/github/issues/Ahmed-Sabri/ymail)](https://github.com/Ahmed-Sabri/ymail/issues)

> **Calculate exact IMAP mailbox storage usage with 100% accuracy — zero estimation, zero error margin.**

`ymail` is a production-ready Python utility that connects to IMAP email servers, analyzes every folder, and computes **exact** storage consumption and message counts. Built for IT administrators, compliance teams, and anyone who needs authoritative mailbox metrics — not estimates.

---

## ✨ Features

### 🔍 Accuracy First

| Method | Description | Accuracy |
|--------|-------------|----------|
| **STATUS=SIZE** (RFC 8438) | Server-provided exact mailbox size in one command | ✅ 100% exact |
| **RFC822.SIZE iteration** | Fetches size metadata for every message (no body download) | ✅ 100% exact |
| **QUOTA integration** | Uses server-authoritative storage limits where available | ✅ Authoritative |
| **No sampling** | Unlike other tools, `ymail` never estimates — every message counts | ✅ Deterministic |

### 🚀 Practical & Scalable

- **Batch processing**: Analyze dozens or hundreds of mailboxes from a single CSV/Excel file
- **Chunked fetching**: Processes messages in configurable chunks to avoid timeouts
- **Rate limiting**: Built-in delays between accounts to respect server policies
- **Robust error handling**: Continues processing other accounts if one fails
- **Detailed logging**: Real-time progress with emoji-enhanced output for readability

### 📊 Rich Output

- **Excel report** with timestamped filename
- Per-account summary: total messages, total size (MB/GB)
- Per-folder breakdown: messages, size, and method used
- Quota information: used/limit/percentage (when supported)
- Column sorting for easy analysis in Excel/LibreOffice

### 🔐 Secure by Design

- Credentials read from input file only — never logged or stored in output
- SSL/TLS connections by default (port 993)
- Read-only IMAP operations — no flags or messages are modified

---

## 📋 Requirements

### System
- Python 3.8 or higher
- Internet connection to target IMAP servers

### Python Dependencies

```bash
pandas>=1.5.0
openpyxl>=3.0.0  # Required for Excel (.xlsx) output
```

Install dependencies:

```bash
pip install pandas openpyxl
```

> 💡 **Note**: `imaplib`, `datetime`, `time`, and `sys` are part of Python's standard library — no installation needed.

---

## 🚀 Installation

### Option 1: Clone from GitHub (Recommended)

```bash
git clone https://github.com/Ahmed-Sabri/ymail.git
cd ymail
pip install pandas openpyxl
```

### Option 2: Direct Download

1. Download [`app.py`](https://github.com/Ahmed-Sabri/ymail/blob/main/app.py) to your working directory
2. Install dependencies manually:

```bash
pip install pandas openpyxl
```

### Option 3: Quick Test (No Install)

```bash
# Clone repo
git clone https://github.com/Ahmed-Sabri/ymail.git
cd ymail

# Install deps
pip install pandas openpyxl

# Run with example data
python app.py
```

---

## 📁 Input File Format

Create a CSV or Excel file with **no header row** and exactly 3 columns:

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Email Address | Password | IMAP Server |

### Example: `mailboxes.csv`

```csv
user1@example.com,SecurePass123,imap.example.com
user2@company.org,MyP@ssw0rd!,imap.mail.yahoo.com
admin@domain.net,AdminKey789,imap.gmail.com
```

### Example: `mailboxes.xlsx`

Same structure, saved as Excel format (`.xlsx`).  
📥 **[Download example file](example_mailboxes.xlsx)** (contains dummy data for testing)

> ⚠️ **Security Note**  
> - Store input files securely and delete after analysis  
> - Consider using app-specific passwords for accounts with 2FA  
> - Never commit real credentials to version control

---

## ▶️ Usage

### Basic Run

```bash
python app.py
```

By default, the script:
1. Reads `mailboxes.xlsx` from the current directory
2. Analyzes each account sequentially
3. Saves results to `mailbox_analysis_YYYYMMDD_HHMMSS.xlsx`

### Custom Input/Output

Edit the `main()` function in `app.py`:

```python
def main():
    analyzer = MailboxAnalyzer()
    
    # Change input file path as needed
    input_file = "path/to/your_accounts.csv"  # or .xlsx
    
    # Optional: customize output filename
    output_file = f"custom_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    analyzer.process_file(input_file, output_file)
```

### Command-Line Override (Advanced)

For flexibility, you can modify the script to accept arguments:

```python
# Add near top of app.py
import argparse

# In main():
parser = argparse.ArgumentParser(description='Analyze IMAP mailbox storage')
parser.add_argument('-i', '--input', default='mailboxes.xlsx', help='Input file path')
parser.add_argument('-o', '--output', help='Output file path (optional)')
args = parser.parse_args()

analyzer.process_file(args.input, args.output or f"mailbox_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
```

---

## 📤 Output Explained

The generated Excel file contains one row per analyzed mailbox with these columns:

### Core Metrics

| Column | Description |
|--------|-------------|
| `email` | Account email address |
| `imap_server` | IMAP server hostname |
| `total_messages` | Exact count of all messages across all folders |
| `total_size_mb` | Exact total storage used (MB, 2 decimal places) |
| `analysis_timestamp` | When the analysis completed |

### Quota Information (if supported by server)

| Column | Description |
|--------|-------------|
| `quota_used_mb` | Server-reported used storage (MB) |
| `quota_limit_mb` | Server-reported quota limit (MB) |
| `quota_usage_percent` | Percentage of quota consumed |

### Per-Folder Metrics

For each folder found (e.g., `INBOX`, `Sent`, `Archive`):

| Column Pattern | Example | Description |
|----------------|---------|-------------|
| `{folder}_messages` | `INBOX_messages` | Exact message count in folder |
| `{folder}_size_mb` | `INBOX_size_mb` | Exact storage used by folder (MB) |
| `{folder}_method` | `INBOX_method` | Method used: `STATUS=SIZE`, `RFC822.SIZE`, or `error` |

> Folder names are sanitized: `/`, spaces, `[`, `]`, `\`, `.` → `_`  
> Example: `Drafts/template` → `Drafts_template_messages`

---

## ⚙️ Configuration Options

Edit these values in `app.py` to tune behavior:

### Performance Tuning

```python
# In get_folder_info() method
chunk_size = 500  # Messages per UID FETCH batch
# ↑ Increase for faster processing (risk: server timeout)
# ↓ Decrease for stability on slow connections

# In process_file() method
time.sleep(2)  # Delay between accounts
# ↑ Increase to be more polite to servers
# ↓ Decrease for faster batch runs (watch for rate limits)
```

### Server Compatibility

```python
# In connect_to_imap()
use_ssl = True   # Default: True (port 993)
port = 993       # Change to 143 if using STARTTLS instead
```

### Debug Mode

Temporarily enable verbose parsing logs:

```python
# In analyze_mailbox(), change:
info = self.get_folder_info(mail, folder_name)
# To:
info = self.get_folder_info(mail, folder_name, debug_parse=True)
```

This prints raw IMAP responses to help troubleshoot server-specific formats.

---

## ⏱️ Performance Expectations

| Scenario | STATUS=SIZE Supported | RFC822.SIZE Fallback |
|----------|----------------------|-------------------|
| **Small folder** (<100 msgs) | ~1 second | ~5-10 seconds |
| **Medium folder** (1K msgs) | ~1 second | ~30-60 seconds |
| **Large folder** (10K msgs) | ~1 second | ~5-10 minutes |
| **Very large** (100K+ msgs) | ~1 second | ~1-2 hours |

### Tips for Large-Scale Runs

1. **Run off-peak**: Analyze mailboxes during low-traffic hours
2. **Parallelize**: Use `multiprocessing` to analyze multiple accounts concurrently (respect per-server rate limits)
3. **Filter folders**: Modify `analyze_mailbox()` to skip non-essential folders (e.g., `Trash`, `Spam`)
4. **Log to file**: Redirect output for unattended runs:

```bash
python app.py > analysis_log.txt 2>&1
```

---

## 🔧 Troubleshooting

### ❌ "Nothing happens when I run the script"

- Ensure `if __name__ == "__main__": main()` exists at the bottom of `app.py`
- Add `print("Script loaded")` after imports to verify execution
- Confirm input file exists in the working directory: `ls -la mailboxes.xlsx`

### ❌ "Connection failed" or "Login error"

- Verify IMAP server hostname and port (common: `imap.gmail.com:993`, `imap.mail.yahoo.com:993`)
- Check if app-specific password is required (Gmail, Yahoo, Outlook)
- Ensure IMAP access is enabled in the email account settings

### ❌ "Could not parse folder names" or "EXAMINE command error"

- Server may use non-standard folder naming — check debug output
- Ensure `_parse_folder_name()` regex handles your server's LIST response format
- Try enabling `debug_parse=True` to inspect raw responses

### ❌ "All sizes show 0.00 MB"

- Server may not support `RFC822.SIZE` parsing in expected format
- Enable debug mode to see raw fetch responses
- Verify `re.search(r'RFC822\.SIZE\s+(\d+)', ...)` matches your server's output

### ❌ Script hangs on large folders

- Reduce `chunk_size` from 500 → 100 in `get_folder_info()`
- Increase timeout: add `mail.socket().settimeout(60)` after login
- Check network stability and server rate limits

---

## 🌐 Supported IMAP Providers

| Provider | STATUS=SIZE | QUOTA | Notes |
|----------|-------------|-------|-------|
| **Yahoo Mail** | ❌ | ❌ | Falls back to RFC822.SIZE iteration |
| **Gmail** | ❌ | ⚠️ Limited | Use app passwords; IMAP must be enabled |
| **Outlook/Office 365** | ⚠️ Varies | ✅ | May require modern auth; test first |
| **cPanel/Dovecot** | ✅ Often | ✅ | Typically full support |
| **Custom Dovecot** | ✅ | ✅ | Configure `plugin { quota = maildir* }` |
| **Zimbra** | ⚠️ | ✅ | Check server capabilities |

> 💡 Run a single-account test first to verify compatibility before batch processing.

---

## 🤝 Contributing

Contributions are welcome! Here's how to help:

1. **Fork** the repository
2. **Create a feature branch**: `git checkout -b feature/amazing-idea`
3. **Commit changes**: `git commit -m 'Add amazing feature'`
4. **Push to branch**: `git push origin feature/amazing-idea`
5. **Open a Pull Request**

### Areas for Improvement

- [ ] Add command-line argument parsing (`argparse`)
- [ ] Support JSON/CSV output formats
- [ ] Add progress bar (`tqdm`) for large runs
- [ ] Implement multiprocessing for parallel account analysis
- [ ] Add unit tests with mock IMAP server
- [ ] Support OAuth2 authentication for modern providers

### Code Style

- Follow [PEP 8](https://peps.python.org/pep-0008/)
- Use type hints for new functions
- Add docstrings to public methods
- Keep functions focused and testable

---

## 📜 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

```text
MIT License

Copyright (c) 2026 Ahmed Sabri

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👤 Author

**Ahmed Sabri**  
🔗 [GitHub](https://github.com/Ahmed-Sabri)  


---

## 🙏 Acknowledgments

- Python `imaplib` documentation — the foundation of IMAP interactions
- RFC 8438 (STATUS=SIZE) and RFC 2087 (QUOTA) — for defining server capabilities
- The open-source community — for inspiration and best practices
- You — for using and improving this tool! 🎉

---

> 🗂️ **Repository**: https://github.com/Ahmed-Sabri/ymail  
> 🐛 **Issues**: https://github.com/Ahmed-Sabri/ymail/issues  
> ✨ **Feature Requests**: https://github.com/Ahmed-Sabri/ymail/discussions

*Built with ❤️ for accurate, transparent, and ethical email infrastructure management.*
