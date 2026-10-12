"""Импорт генератора с существующим CLI-именем; исходный manifest единственный."""
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("section_recipe_api", Path(__file__).with_name("Build-SectionRecipeApi.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
ROOT, DOCS, inventory = module.ROOT, module.DOCS, module.inventory
