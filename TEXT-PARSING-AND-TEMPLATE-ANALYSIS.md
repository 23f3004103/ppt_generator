# Text Parsing and Template Style Application - Technical Explanation

## Input Text Parsing and Slide Mapping Process

Our PowerPoint generator employs a sophisticated multi-stage approach to transform unstructured text into coherent slide content. The process begins with **semantic analysis** using Large Language Models (specifically NVIDIA's Llama 3.3 70B) to identify natural break points and thematic boundaries within the input text. Rather than applying arbitrary slide limits, the system calculates optimal slide count based on content density, word count, and user guidance parameters.

The **intelligent mapping algorithm** first segments the text into logical sections by analyzing sentence structures and paragraph transitions. Each segment is then evaluated for slide-worthiness using topic clustering techniques that group related concepts together. The LLM API processes this segmented content to extract key points, supporting details, and hierarchical relationships, ensuring each slide maintains focused messaging while contributing to the overall narrative flow.

**Content structure recognition** adapts to different presentation types: investor pitches follow Problem→Solution→Market→Traction sequences, while training materials use Objectives→Content→Practice→Summary patterns. This contextual awareness ensures that generated slides follow professional presentation conventions.

## Template Style and Asset Application

Visual consistency is achieved through **comprehensive template analysis** that occurs entirely client-side using browser-based zip extraction and XML parsing. When users upload a PowerPoint template, our system extracts the underlying structure including slide layouts, color schemes, font definitions, and embedded media assets.

The **style transfer mechanism** maps template elements to generated content through intelligent matching algorithms. Master slide layouts are analyzed to identify content placeholders, and our system selects the most appropriate layout for each slide type (title, content, comparison, etc.). Original color palettes, typography settings, and design elements are preserved by copying XML style definitions directly from the template.

**Asset reuse** is handled by extracting images, logos, and graphics from the template's media folder and strategically placing them based on content context and layout compatibility. This approach ensures that generated presentations maintain the visual identity and professional appearance of the original template while accommodating new content structure and requirements.

---

**Implementation**: Client-side processing ensures privacy while PptxGenJS library handles PowerPoint generation with full template compatibility.
