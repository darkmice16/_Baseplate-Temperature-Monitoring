# Base Plate Monitoring System

A Python-based system for monitoring base plate temperatures using GPIB instrumentation.

## Description

This system monitors temperature data from instrumentation connected via GPIB interface and provides comprehensive logging capabilities. It can also process video data for monitoring cooling systems such as fans, allowing for complete thermal management solutions.

## Features

- Real-time temperature monitoring via GPIB interface
- Automated data logging with timestamped entries
- Video monitoring capabilities for cooling system verification
- Configurable alert thresholds
- Historical data analysis
- Long-term logging support

## Project Structure

```
├── Base Plate Monitoring System.py  # Main application file
├── config.ini                       # Configuration file (not included in repo)
├── logs/                            # Temperature log files
│   └── temperature_monitor_*.log    # Daily temperature logs
└── videos/                          # Video monitoring files
    ├── rotating_fan.mp4             # Example of functioning cooling
    └── stopped_fan.mp4              # Example of failed cooling
```

## Installation

1. Clone this repository
2. Install required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Copy `config.example.ini` to `config.ini` and update with your settings

## Usage

Run the main script:

```
python "Base Plate Monitoring System.py"
```

The system will:
1. Connect to the configured GPIB devices
2. Begin monitoring temperatures at the specified intervals
3. Log data to timestamped files in the logs directory
4. Alert if temperatures exceed configured thresholds

## Configuration

The system uses a configuration file (`config.ini`) for settings such as:
- GPIB device addresses and communication parameters
- Sampling intervals
- Temperature thresholds for alerts
- Log file locations and rotation settings
- Video monitoring configuration

## Requirements

- Python 3.6+
- GPIB interface hardware
- Required Python packages (see requirements.txt)

## License

[MIT License](LICENSE)

## Contributing

Contributions to improve the Base Plate Monitoring System are welcome. Please feel free to submit a Pull Request.

## Contact

For questions and support, please open an issue in the GitHub repository.