from pptx import Presentation
import pandas as pd

# Crear una nueva presentación
prs = Presentation()

# Imprimir los diseños disponibles
print("Diseños disponibles:")
for idx, layout in enumerate(prs.slide_layouts):
    print(idx, layout.name)
