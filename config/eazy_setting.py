import json,sys
from pathlib import Path 

def load_config():

    if getattr(sys,'frozen',False):
        config_path=Path(sys.executable).parent / "config.json"
    else:
        config_path=Path(__file__).parent / "config.json"

    DEFAULT_CONFIG={
        "head_pos":1,
        "head_pos_end":3,
        "head_new_pos":4
        }
    
    if not config_path.exists():
        config_path.write_text(json.dumps(DEFAULT_CONFIG,indent=4))
        return DEFAULT_CONFIG

    with open(config_path,"r") as f:

        try:
            return json.load(f)
        except json.JSONDecodeError:
            print("配置文件损坏，重置默认值")
            config_path.write_text(json.dumps(DEFAULT_CONFIG,indent=4))

