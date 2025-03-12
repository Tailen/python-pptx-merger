import os
import pathlib
from dataclasses import dataclass
from io import BytesIO

import pythonnet
pythonnet.load("coreclr")
import clr
module_dir = pathlib.Path(__file__).parent.absolute()
dll_path = os.path.join(module_dir, "lib", "OpenXmlPowerTools")
clr.AddReference(dll_path)
from System.Collections import Generic  # type: ignore
from OpenXmlPowerTools import (  # type: ignore
    PmlDocument,
    SlideSource,
    PresentationBuilder,
)


@dataclass
class SlideRef:
    """
    A reference to a slide in a source document.
    """
    doc_idx: int
    slide_idx: int


class Merger:
    def merge_slides(
        self,
        src_docs: list[BytesIO],
        slide_refs: list[SlideRef] = [],
    ) -> BytesIO:
        """
        Merge slides from src_docs and returns the merged document.
        slide_refs is an ordered list of slide references to merge.
        If empty, all slides from all src_docs will be merged.
        """
        # Load PmlDocuments
        pml_docs = []
        for i, doc_data in enumerate(src_docs):
            pml_docs.append(PmlDocument(f"document_{i}.pptx", doc_data.getvalue()))

        sources = Generic.List[SlideSource]()
        if not slide_refs:
            # If no specific slides are requested, include all slides from all documents
            for doc in pml_docs:
                sources.Add(SlideSource(doc, True))
        else:
            # Only group consecutive slides from the same document
            current_doc_idx = -1
            start_slide_idx = -1
            count = 0

            for i, ref in enumerate(slide_refs):
                if ref.doc_idx != current_doc_idx:
                    # Add previous range if it exists
                    if current_doc_idx >= 0 and count > 0:
                        sources.Add(
                            SlideSource(
                                pml_docs[current_doc_idx],
                                start_slide_idx,
                                count,
                                True,
                            )
                        )
                    # Start new range
                    current_doc_idx = ref.doc_idx
                    start_slide_idx = ref.slide_idx
                    count = 1
                elif ref.slide_idx == start_slide_idx + count:
                    # Continue current range - consecutive slide in same document
                    count += 1
                else:
                    # Add previous range and start new one in same document
                    sources.Add(
                        SlideSource(
                            pml_docs[current_doc_idx],
                            start_slide_idx,
                            count,
                            True,
                        )
                    )
                    start_slide_idx = ref.slide_idx
                    count = 1

            # Add the final range
            if count > 0:
                sources.Add(
                    SlideSource(
                        pml_docs[current_doc_idx],
                        start_slide_idx,
                        count,
                        True,
                    )
                )

        # Build the presentation
        result_doc = PresentationBuilder.BuildPresentation(sources)
        return BytesIO(result_doc.DocumentByteArray)
