setup(
    name="excel-bom-compare",
    version="1.0.0",
    description="Script to compare piping isometric excel BOMs from SP3D",
    long_description=README,
    long_description_content_type="text/markdown",
    url="https://github.com/jeffcall-ch/excel_bom_compare",
    author="Laszlo Sziva",
    author_email="szivalaszlo@gmail.com",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
    ],
    entry_points={"console_scripts": ["realpython=calculate_difference_bom:main"]},
)
