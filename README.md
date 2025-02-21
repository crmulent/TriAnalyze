# TriAnalyze
a powerful forensic analysis tool that consolidates outputs from ExifTool, OleDump, and NetworkMiner, providing a comprehensive view of digital artifacts. It seamlessly integrates metadata, file structures, and network data to assist in efficient investigations and data analysis.

## Installation

### Clone the repository:
```bash
git clone https://github.com/crmulent/TriAnalyze
cd TriAnalyze
```

### Install the required Python dependencies
```bash
pip install -r requirements.txt
```

## Usage
```bash
python TriAnalyze.py -p <path_to_pcap_file> -f <output_format> -o <output_directory>
```