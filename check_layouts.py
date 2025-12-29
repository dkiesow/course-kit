from pptx import Presentation
import zipfile, shutil, os

# Copy and convert POTX to PPTX
shutil.copy2('4734_template.potx', 'temp_check.pptx')
z = zipfile.ZipFile('temp_check.pptx', 'r')
files = {n: z.read(n) for n in z.namelist()}
z.close()

# Fix content type
content_xml = files['[Content_Types].xml'].decode('utf-8')
content_xml = content_xml.replace('presentationml.template.main', 'presentationml.presentation.main')
files['[Content_Types].xml'] = content_xml.encode('utf-8')

z = zipfile.ZipFile('temp_check.pptx', 'w')
for n, d in files.items():
    z.writestr(n, d)
z.close()

# Load and print layouts with master slide info
prs = Presentation('temp_check.pptx')
print("Slide Masters and Layouts in 4734_template.potx:")
print()

# Build complete layout list from all masters
all_layouts = []
for master_idx, master in enumerate(prs.slide_masters):
    print(f"Master {master_idx}: {master.name if hasattr(master, 'name') else 'Unnamed'}")
    print(f"  Child layouts ({len(master.slide_layouts)}):")
    
    for layout_idx, layout in enumerate(master.slide_layouts):
        print(f"    {layout_idx}: {layout.name}")
        all_layouts.append(layout)
    print()

print(f"\nAll layouts across all masters (total: {len(all_layouts)}):")
for i, layout in enumerate(all_layouts):
    print(f'{i}: {layout.name}')

print(f"\nprs.slide_layouts returns {len(prs.slide_layouts)} layouts")

# Cleanup
os.remove('temp_check.pptx')
