from setuptools import setup


setup(
    name="summarize-ppt-notes",
    version="0.4.0",
    description="Codex Skill and CLI for turning slide decks into detailed Word study notes.",
    packages=["ppt_notes_summarizer", "scripts"],
    python_requires=">=3.9",
    entry_points={
        "console_scripts": [
            "ppt-notes-exporter=ppt_notes_summarizer.cli:main",
            "ppt-notes-doctor=ppt_notes_summarizer.doctor:main",
        ]
    },
    extras_require={
        "render": ["PyMuPDF>=1.24", "Pillow>=10"],
        "dev": ["PyYAML>=6", "pytest>=8"],
    },
)
