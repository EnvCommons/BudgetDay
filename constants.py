from pathlib import Path
import os

# Local development (uncomment for dev)
# ENV_PATH = Path(__file__).parent

# Production (deployment)
if os.path.exists("/orwd_data"):
    ENV_PATH = Path("/orwd_data")
else:
    ENV_PATH = Path(__file__).parent