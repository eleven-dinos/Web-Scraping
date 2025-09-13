# Data Scraping Project

A comprehensive data scraping solution for extracting academic and client information using Python automation tools.

# Project Overview

This project consists of two main scraping components:

1. **Web_Scraping.ipynb** - FAST NUCES professors data scraper across multiple campuses
2. **Aesthetics_Pro.py** - PDF data extraction tool for 5000+ client files

# Features

## Web Scraping (FAST NUCES Professors)
- Scrapes professor data from multiple FAST NUCES campuses
- Uses BeautifulSoup for HTML parsing
- Extracts comprehensive faculty information
- Jupyter notebook format for interactive data exploration

## Aesthetics Pro PDF Scraper
- Processes 5000+ client PDF files
- Automated login system with environment variable security
- Selenium WebDriver for browser automation
- PyAutoGUI for additional UI automation
- Bulk PDF data extraction capabilities

# Technologies Used

- **Python 3.x**
- **Selenium WebDriver** - Browser automation
- **BeautifulSoup4** - HTML parsing and web scraping
- **PyAutoGUI** - GUI automation
- **Jupyter Notebook** - Interactive development environment
- **OS Environment Variables** - Secure credential management

# Installation

## Prerequisites
- Python 3.7 or higher
- Chrome/Firefox browser
- ChromeDriver/GeckoDriver

## Install Dependencies

```bash
pip install selenium beautifulsoup4 pyautogui jupyter pandas requests lxml
```

## WebDriver Setup
1. Download ChromeDriver from [https://chromedriver.chromium.org/](https://chromedriver.chromium.org/)
2. Add ChromeDriver to your system PATH
3. Ensure Chrome browser is installed

# Configuration

## Environment Variables
Create a `.env` file in the project root directory:

```env
USERNAME=your_username_here
PASSWORD=your_password_here
```

## Security Note
- Never commit credentials to version control
- Use environment variables for sensitive information
- Add `.env` to your `.gitignore` file

# Usage

## FAST NUCES Professor Data Scraping

```bash
jupyter notebook Web_Scraping.ipynb
```

Run all cells in the notebook to:
- Connect to FAST NUCES websites
- Extract professor information
- Parse and clean data
- Export results to CSV/Excel

## Aesthetics Pro PDF Scraper

```bash
python Aesthetics_Pro.py
```

The script will:
- Load credentials from environment variables
- Automate browser login
- Navigate to letter for which client needs to process
- Navigate to client page
- Navigate through PDF files
- Extract and process client data
- Save results to specified output format

# Project Structure

```
├── Web_Scraping.ipynb          # FAST NUCES professors scraper
├── Aesthetics_Pro.py           # PDF client data scraper
├── .env                        # Environment variables (not in repo)
├── .gitignore                  # Git ignore file
├── README.md                   # Project documentation
├── requirements.txt            # Python dependencies
```

# Configuration Options

## Web Scraping Settings
- Campus selection

## PDF Scraper Settings
- Letter to start with
- Client number to start with
- File path configurations
- Error handling preferences

# Important Notes

## Rate Limiting
- Implement appropriate delays between requests
- Respect website robots.txt files
- Avoid overwhelming target servers

## Legal Considerations
- Ensure compliance with website terms of service
- Respect data privacy regulations
- Use scraped data responsibly

## Error Handling
- The scripts include robust error handling
- Check logs for debugging information
- Ensure stable internet connection

# Troubleshooting

## Common Issues

**WebDriver Errors:**
- Ensure ChromeDriver version matches Chrome browser
- Check PATH configuration
- Verify WebDriver permissions

**Authentication Failures:**
- Verify environment variables are set correctly
- Check username/password validity
- Ensure login page elements haven't changed

**PDF Processing Issues:**
- Verify PDF file permissions
- Check available disk space
- Ensure PyAutoGUI has necessary permissions

# Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

# License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

# Disclaimer

This tool is for educational and research purposes. Users are responsible for ensuring compliance with all applicable laws and website terms of service. The authors are not responsible for any misuse of this software.

# Support

For questions, issues, or contributions, please:
- Open an issue in the GitHub repository
- Check existing documentation
- Review troubleshooting section

---

**Note:** Always ensure you have proper authorization before scraping any website or processing client data. Respect privacy laws and data protection regulations in your jurisdiction.
