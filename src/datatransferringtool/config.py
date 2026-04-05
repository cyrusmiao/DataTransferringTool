import yaml
from pathlib import Path
from dataclasses import dataclass
from typing import List, Literal, Dict

@dataclass
class SourceConfig:
    file_path: str
    reference_column: Dict[str, str]
    mapping: Dict[str, str]

@dataclass
class TransferConfig:
    target_file: str
    output_file: str
    conflict_resolution: Literal['keep_original', 'overwrite', 'manual']
    sources: List[SourceConfig]

def load_config(yaml_path: str | Path) -> TransferConfig:
    with open(yaml_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    
    sources = []
    for src in data.get('sources', []):
        sources.append(SourceConfig(
            file_path=src['file_path'],
            reference_column=src.get('reference_column', {}),
            mapping=src.get('mapping', {})
        ))
        
    return TransferConfig(
        target_file=data['target_file'],
        output_file=data.get('output_file', 'output.xlsx'),
        conflict_resolution=data.get('conflict_resolution', 'keep_original'),
        sources=sources
    )
