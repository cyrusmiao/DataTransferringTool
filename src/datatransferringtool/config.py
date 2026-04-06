import yaml
from pathlib import Path
from dataclasses import dataclass
from typing import List, Literal, Dict


class MappingPairs(list):
    pass


class PreserveMappingLoader(yaml.SafeLoader):
    pass


def _construct_mapping_pairs(loader, node, deep=False):
    pairs = MappingPairs()
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        value = loader.construct_object(value_node, deep=deep)
        pairs.append((key, value))
    return pairs


PreserveMappingLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
    _construct_mapping_pairs
)

@dataclass
class SourceConfig:
    file_path: str
    sheet_name: str | int | None
    reference_column: Dict[str, str]
    mapping: List[tuple[str, str]]

@dataclass
class TransferConfig:
    target_file: str
    target_sheet: str | int | None
    output_file: str
    generate_transfer_report: bool
    generate_reference_report: bool
    highlight_conflict_cells: bool
    conflict_resolution: Literal['keep_original', 'overwrite', 'manual']
    sources: List[SourceConfig]


def _normalize_mapping(raw_mapping) -> List[tuple[str, str]]:
    if raw_mapping is None:
        return []
    if isinstance(raw_mapping, MappingPairs):
        return [(str(src), str(tgt)) for src, tgt in raw_mapping]
    if isinstance(raw_mapping, list):
        mapping = []
        for item in raw_mapping:
            if isinstance(item, MappingPairs):
                mapping.extend((str(src), str(tgt)) for src, tgt in item)
        return mapping
    return [(str(src), str(tgt)) for src, tgt in dict(raw_mapping).items()]


def _normalize_reference_column(raw_reference_column) -> Dict[str, str]:
    if raw_reference_column is None:
        return {}
    if isinstance(raw_reference_column, MappingPairs):
        return {str(src): str(tgt) for src, tgt in raw_reference_column}
    if isinstance(raw_reference_column, list):
        reference_column = {}
        for item in raw_reference_column:
            if isinstance(item, MappingPairs):
                reference_column.update((str(src), str(tgt)) for src, tgt in item)
        return reference_column
    return {str(src): str(tgt) for src, tgt in dict(raw_reference_column).items()}

def load_config(yaml_path: str | Path) -> TransferConfig:
    with open(yaml_path, 'r', encoding='utf-8') as f:
        data = yaml.load(f, Loader=PreserveMappingLoader)

    data = {key: value for key, value in data}
    
    sources = []
    for raw_src in data.get('sources', []):
        src = {key: value for key, value in raw_src}
        sources.append(SourceConfig(
            file_path=src['file_path'],
            sheet_name=src.get('sheet_name'),
            reference_column=_normalize_reference_column(src.get('reference_column')),
            mapping=_normalize_mapping(src.get('mapping'))
        ))
        
    return TransferConfig(
        target_file=data['target_file'],
        target_sheet=data.get('target_sheet'),
        output_file=data.get('output_file', 'output.xlsx'),
        generate_transfer_report=data.get('generate_transfer_report', data.get('generate_report', False)),
        generate_reference_report=data.get('generate_reference_report', False),
        highlight_conflict_cells=data.get('highlight_conflict_cells', False),
        conflict_resolution=data.get('conflict_resolution', 'keep_original'),
        sources=sources
    )
