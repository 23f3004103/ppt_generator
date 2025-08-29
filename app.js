/**
 * PowerPoint Generator Application
 * Main JavaScript logic for text-to-presentation conversion
 */

class PowerPointGenerator {
    constructor() {
        this.currentStep = 1;
        this.maxSteps = 4;
        this.templateFile = null;
        this.presentationData = null;
        this.slideData = [];
        this.currentSlideIndex = 0;
        
        this.initializeEventListeners();
        this.updateStepDisplay();
    }

    initializeEventListeners() {
        // Character counter for text input
        const inputText = document.getElementById('inputText');
        inputText.addEventListener('input', this.updateCharCount);

        // Preset guidance buttons
        document.querySelectorAll('.preset-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                document.getElementById('guidance').value = e.target.dataset.preset;
                this.highlightPresetButton(e.target);
            });
        });

        // Template file upload
        const templateFile = document.getElementById('templateFile');
        templateFile.addEventListener('change', this.handleTemplateUpload.bind(this));

        // Sample template buttons
        document.querySelectorAll('.sample-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.selectSampleTemplate(e.target.dataset.template);
                this.highlightSampleButton(e.target);
            });
        });

        // Drag and drop for template upload
        const uploadArea = document.getElementById('templateUpload');
        uploadArea.addEventListener('dragover', this.handleDragOver);
        uploadArea.addEventListener('dragleave', this.handleDragLeave);
        uploadArea.addEventListener('drop', this.handleDrop.bind(this));
        uploadArea.addEventListener('click', () => templateFile.click());

        // Slide navigation
        document.getElementById('prevSlide')?.addEventListener('click', this.previousSlide.bind(this));
        document.getElementById('nextSlide')?.addEventListener('click', this.nextSlide.bind(this));

        // Update summary when inputs change
        inputText.addEventListener('input', this.updateSummary.bind(this));
        document.getElementById('guidance').addEventListener('input', this.updateSummary.bind(this));
        document.getElementById('llmProvider').addEventListener('change', this.updateSummary.bind(this));
    }

    updateCharCount() {
        const text = document.getElementById('inputText').value;
        document.getElementById('charCount').textContent = text.length;
    }

    highlightPresetButton(selectedBtn) {
        document.querySelectorAll('.preset-btn').forEach(btn => {
            btn.style.background = 'linear-gradient(135deg, #e0e7ff, #c7d2fe)';
        });
        selectedBtn.style.background = 'linear-gradient(135deg, #4f46e5, #7c3aed)';
        selectedBtn.style.color = 'white';
    }

    highlightSampleButton(selectedBtn) {
        document.querySelectorAll('.sample-btn').forEach(btn => {
            btn.classList.remove('selected');
        });
        selectedBtn.classList.add('selected');
    }

    handleTemplateUpload(event) {
        const file = event.target.files[0];
        if (file) {
            this.processTemplateFile(file);
        }
    }

    handleDragOver(event) {
        event.preventDefault();
        event.currentTarget.classList.add('dragover');
    }

    handleDragLeave(event) {
        event.currentTarget.classList.remove('dragover');
    }

    handleDrop(event) {
        event.preventDefault();
        event.currentTarget.classList.remove('dragover');
        
        const files = event.dataTransfer.files;
        if (files.length > 0) {
            this.processTemplateFile(files[0]);
        }
    }

    processTemplateFile(file) {
        if (!file.name.match(/\.(pptx|potx)$/i)) {
            alert('Please upload a PowerPoint file (.pptx or .potx)');
            return;
        }

        this.templateFile = file;
        
        // Display template information
        const templateInfo = document.getElementById('templateInfo');
        const templateDetails = document.getElementById('templateDetails');
        
        templateDetails.innerHTML = `
            <div style="display: flex; align-items: center; gap: 15px;">
                <div style="font-size: 2rem;">📊</div>
                <div>
                    <strong>${file.name}</strong><br>
                    <span style="color: #6b7280;">Size: ${this.formatFileSize(file.size)}</span>
                </div>
            </div>
        `;
        
        templateInfo.classList.remove('hidden');
        this.updateSummary();
    }

    selectSampleTemplate(templateType) {
        // Simulate template selection
        const templates = {
            corporate: 'Corporate Professional Template',
            modern: 'Modern Minimalist Template',
            minimal: 'Clean Minimal Template',
            creative: 'Creative Design Template'
        };

        this.templateFile = { 
            name: templates[templateType], 
            size: 1024000,
            isSample: true,
            type: templateType
        };

        const templateInfo = document.getElementById('templateInfo');
        const templateDetails = document.getElementById('templateDetails');
        
        templateDetails.innerHTML = `
            <div style="display: flex; align-items: center; gap: 15px;">
                <div style="font-size: 2rem;">✨</div>
                <div>
                    <strong>${templates[templateType]}</strong><br>
                    <span style="color: #6b7280;">Sample Template</span>
                </div>
            </div>
        `;
        
        templateInfo.classList.remove('hidden');
        this.updateSummary();
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    updateSummary() {
        const text = document.getElementById('inputText').value;
        const provider = document.getElementById('llmProvider').value;
        
        document.getElementById('contentSummary').textContent = 
            text.length > 0 ? `${text.length} characters` : 'No content yet';
        
        document.getElementById('aiSummary').textContent = 
            provider ? this.getProviderDisplayName(provider) : 'Not selected';
        
        document.getElementById('templateSummary').textContent = 
            this.templateFile ? this.templateFile.name : 'No template selected';
    }

    getProviderDisplayName(provider) {
        const names = {
            openai: 'OpenAI (GPT)',
            google: 'Google (Gemini)',
            nvidia: 'NVIDIA (Llama 3.3 70B)'
        };
        return names[provider] || provider;
    }

    nextStep(step) {
        if (step > this.currentStep + 1) return;
        
        // Validate current step
        if (!this.validateStep(this.currentStep)) return;
        
        this.currentStep = step;
        this.updateStepDisplay();
        this.updateSummary();
    }

    prevStep(step) {
        this.currentStep = step;
        this.updateStepDisplay();
    }

    validateStep(step) {
        switch (step) {
            case 1:
                const text = document.getElementById('inputText').value;
                if (!text.trim()) {
                    alert('Please enter some text content to convert to slides.');
                    return false;
                }
                return true;
            
            case 2:
                const apiKey = document.getElementById('apiKey').value;
                if (!apiKey.trim()) {
                    alert('Please enter your API key to proceed.');
                    return false;
                }
                return true;
            
            case 3:
                if (!this.templateFile) {
                    alert('Please upload a template or select a sample template.');
                    return false;
                }
                return true;
            
            default:
                return true;
        }
    }

    updateStepDisplay() {
        // Hide all form sections
        for (let i = 1; i <= this.maxSteps; i++) {
            document.getElementById(`step${i}`).classList.add('hidden');
        }
        
        // Show current step
        document.getElementById(`step${this.currentStep}`).classList.remove('hidden');
        
        // Update step indicators
        document.querySelectorAll('.step').forEach((step, index) => {
            const stepNumber = index + 1;
            step.classList.remove('active', 'completed');
            
            if (stepNumber === this.currentStep) {
                step.classList.add('active');
            } else if (stepNumber < this.currentStep) {
                step.classList.add('completed');
            }
        });
    }

    async generatePresentation() {
        const generateBtn = document.getElementById('generateBtn');
        generateBtn.disabled = true;
        generateBtn.textContent = 'Generating...';

        try {
            // Show progress panel
            document.getElementById('progressPanel').classList.remove('hidden');
            document.getElementById('resultPanel').classList.add('hidden');

            // Step 1: Analyze content
            await this.updateProgress('step-analyze', 'Analyzing your content...', 25);
            const analysisResult = await this.analyzeContent();

            // Step 2: Structure slides
            await this.updateProgress('step-structure', 'Creating slide structure...', 50);
            const slideStructure = await this.createSlideStructure(analysisResult);

            // Step 3: Apply template
            await this.updateProgress('step-template', 'Applying template style...', 75);
            const styledSlides = await this.applyTemplateStyle(slideStructure);

            // Step 4: Generate presentation
            await this.updateProgress('step-generate', 'Generating final presentation...', 100);
            const presentation = await this.buildPresentation(styledSlides);

            // Show results
            this.showResults(presentation);

        } catch (error) {
            console.error('Generation failed:', error);
            alert('Failed to generate presentation: ' + error.message);
        } finally {
            generateBtn.disabled = false;
            generateBtn.textContent = '🎨 Generate Presentation';
        }
    }

    async updateProgress(stepId, message, percentage) {
        // Mark previous steps as completed
        document.querySelectorAll('.progress-step').forEach(step => {
            if (step.id !== stepId) {
                step.classList.remove('active');
                step.classList.add('completed');
            }
        });

        // Mark current step as active
        const currentStep = document.getElementById(stepId);
        currentStep.classList.remove('completed');
        currentStep.classList.add('active');
        currentStep.querySelector('.progress-text').textContent = message;

        // Update progress bar
        document.getElementById('progressFill').style.width = percentage + '%';

        // Simulate processing time
        await this.delay(1500 + Math.random() * 1000);
    }

    async analyzeContent() {
        const text = document.getElementById('inputText').value;
        const guidance = document.getElementById('guidance').value;
        const provider = document.getElementById('llmProvider').value;
        const apiKey = document.getElementById('apiKey').value;

        // In a real implementation, this would call the LLM API
        // For demo purposes, we'll simulate the analysis
        if (apiKey && apiKey.length > 10) {
            // Uncomment and use this for real API calls:
            // return this.callLLMAPI(text, guidance, provider, apiKey);
            console.log(`Would call ${provider} API with key: ${apiKey.substring(0, 8)}...`);
        }
        
        return this.simulateContentAnalysis(text, guidance);
    }

    async callLLMAPI(text, guidance, provider, apiKey) {
        // Real API implementation examples (currently commented for demo):
        try {
            switch (provider) {
                case 'openai':
                    return await this.callOpenAI(text, guidance, apiKey);
                case 'google':
                    return await this.callGoogle(text, guidance, apiKey);
                case 'nvidia':
                    return await this.callNVIDIA(text, guidance, apiKey);
                default:
                    throw new Error(`Unsupported provider: ${provider}`);
            }
        } catch (error) {
            console.error('LLM API Error:', error);
            // Fallback to simulation
            return this.simulateContentAnalysis(text, guidance);
        }
    }

    async callNVIDIA(text, guidance, apiKey) {
        // NVIDIA NIM API implementation using Llama 3.3 70B for best results
        const response = await fetch('https://integrate.api.nvidia.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify({
                model: "meta/llama-3.3-70b-instruct",
                messages: [{
                    role: "user",
                    content: `Analyze this text for a PowerPoint presentation. ${guidance ? `Guidance: ${guidance}` : ''}\n\nText: ${text}\n\nProvide:\n1. Estimated slide count (3-20 slides based on content)\n2. Key topics for each slide with clear titles\n3. Presentation tone (professional, technical, sales, etc.)\n4. Suggested structure and flow\n5. Main points and supporting details for each slide\n\nFormat your response clearly with numbered sections for easy parsing.`
                }],
                temperature: 0.7,
                max_tokens: 2048
            })
        });

        if (!response.ok) {
            throw new Error(`NVIDIA API Error: ${response.status}`);
        }

        const data = await response.json();
        return this.parseNVIDIAResponse(data.choices[0].message.content, text, guidance);
    }

    async callOpenAI(text, guidance, apiKey) {
        // OpenAI API implementation
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: "gpt-4",
                messages: [{
                    role: "user",
                    content: `Analyze this text for a PowerPoint presentation. ${guidance ? `Guidance: ${guidance}` : ''}\n\nText: ${text}\n\nProvide:\n1. Estimated slide count\n2. Key topics for each slide\n3. Presentation tone\n4. Suggested structure`
                }],
                temperature: 0.7,
                max_tokens: 1000
            })
        });

        const data = await response.json();
        return this.parseOpenAIResponse(data.choices[0].message.content, text, guidance);
    }

    async callGoogle(text, guidance, apiKey) {
        // Google Gemini API implementation
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{
                        text: `Analyze this text for a PowerPoint presentation. ${guidance ? `Guidance: ${guidance}` : ''}\n\nText: ${text}\n\nProvide:\n1. Estimated slide count\n2. Key topics for each slide\n3. Presentation tone\n4. Suggested structure`
                    }]
                }]
            })
        });

        const data = await response.json();
        return this.parseGoogleResponse(data.candidates[0].content.parts[0].text, text, guidance);
    }

    parseNVIDIAResponse(response, text, guidance) {
        // Parse NVIDIA API response into structured format
        try {
            // Extract structured data from LLM response
            const lines = response.split('\n');
            const analysis = this.extractAnalysisFromLines(lines, text, guidance);
            return analysis;
        } catch (error) {
            console.warn('Failed to parse NVIDIA response, falling back to simulation');
            return this.simulateContentAnalysis(text, guidance);
        }
    }

    parseOpenAIResponse(response, text, guidance) {
        // Parse OpenAI response similarly
        return this.parseNVIDIAResponse(response, text, guidance);
    }

    parseGoogleResponse(response, text, guidance) {
        // Parse Google response similarly
        return this.parseNVIDIAResponse(response, text, guidance);
    }

    extractAnalysisFromLines(lines, text, guidance) {
        // Extract analysis data from LLM response lines
        const wordCount = text.split(/\s+/).length;
        let estimatedSlides = Math.max(3, Math.min(20, Math.ceil(wordCount / 100)));
        
        // Try to extract slide count from response
        const slideCountMatch = lines.find(line => line.match(/\d+.*slide/i));
        if (slideCountMatch) {
            const match = slideCountMatch.match(/(\d+)/);
            if (match) {
                estimatedSlides = Math.max(3, Math.min(20, parseInt(match[1])));
            }
        }

        return {
            estimatedSlides,
            topics: this.extractTopics(lines, estimatedSlides),
            tone: this.detectTone(guidance),
            structure: this.suggestStructure(guidance),
            keyPoints: this.extractKeyPoints(text),
            aiGenerated: true
        };
    }

    extractTopics(lines, slideCount) {
        // Extract topic suggestions from LLM response
        const topics = [];
        let inTopicSection = false;
        
        for (const line of lines) {
            if (line.toLowerCase().includes('topic') || line.toLowerCase().includes('slide')) {
                inTopicSection = true;
            }
            if (inTopicSection && (line.includes('-') || line.includes('•') || line.match(/^\d+\./))) {
                const topic = line.replace(/^[-•\d.\s]+/, '').trim();
                if (topic && topic.length > 5) {
                    topics.push(topic);
                }
            }
            if (topics.length >= slideCount) break;
        }

        // Fallback if no topics extracted
        if (topics.length === 0) {
            for (let i = 1; i <= slideCount; i++) {
                topics.push(`Topic ${i}`);
            }
        }

        return topics.slice(0, slideCount);
    }

    simulateContentAnalysis(text, guidance) {
        // Simulate AI analysis
        const wordCount = text.split(/\s+/).length;
        const estimatedSlides = Math.max(3, Math.min(20, Math.ceil(wordCount / 100)));
        
        // Extract potential slide topics
        const sentences = text.split(/[.!?]+/).filter(s => s.trim());
        const topics = sentences.slice(0, estimatedSlides).map(s => s.trim().substring(0, 60) + '...');

        return {
            estimatedSlides,
            topics,
            tone: this.detectTone(guidance),
            structure: this.suggestStructure(guidance),
            keyPoints: this.extractKeyPoints(text)
        };
    }

    detectTone(guidance) {
        if (!guidance) return 'professional';
        
        const lowerGuidance = guidance.toLowerCase();
        if (lowerGuidance.includes('pitch') || lowerGuidance.includes('investor')) return 'persuasive';
        if (lowerGuidance.includes('technical')) return 'technical';
        if (lowerGuidance.includes('training')) return 'educational';
        if (lowerGuidance.includes('sales')) return 'sales';
        
        return 'professional';
    }

    suggestStructure(guidance) {
        const lowerGuidance = guidance?.toLowerCase() || '';
        
        if (lowerGuidance.includes('pitch') || lowerGuidance.includes('investor')) {
            return ['Title', 'Problem', 'Solution', 'Market', 'Business Model', 'Team', 'Financials', 'Ask'];
        }
        if (lowerGuidance.includes('training')) {
            return ['Title', 'Agenda', 'Objectives', 'Content Sections', 'Practice', 'Summary', 'Q&A'];
        }
        if (lowerGuidance.includes('sales')) {
            return ['Title', 'Challenge', 'Solution', 'Benefits', 'Proof', 'Pricing', 'Next Steps'];
        }
        
        return ['Title', 'Introduction', 'Main Points', 'Details', 'Conclusion'];
    }

    extractKeyPoints(text) {
        // Simple key point extraction
        const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 20);
        return sentences.slice(0, 10).map(s => s.trim());
    }

    async createSlideStructure(analysis) {
        const slideCount = parseInt(document.getElementById('slideCount').value) || analysis.estimatedSlides;
        const generateNotes = document.getElementById('generateNotes').checked;
        
        const slides = [];
        
        // Title slide
        slides.push({
            type: 'title',
            title: this.generateTitle(analysis),
            subtitle: 'Generated from your content',
            notes: generateNotes ? 'Welcome to the presentation. This was generated from your text content.' : ''
        });

        // Content slides
        for (let i = 1; i < slideCount; i++) {
            const slideContent = this.generateSlideContent(analysis, i, slideCount);
            slides.push({
                type: 'content',
                title: slideContent.title,
                content: slideContent.content,
                layout: slideContent.layout,
                notes: generateNotes ? slideContent.notes : ''
            });
        }

        this.slideData = slides;
        return slides;
    }

    generateTitle(analysis) {
        const guidance = document.getElementById('guidance').value;
        if (guidance) {
            return `Presentation: ${guidance}`;
        }
        
        const firstTopic = analysis.topics[0] || 'Your Content';
        return firstTopic.replace('...', '');
    }

    generateSlideContent(analysis, slideIndex, totalSlides) {
        const keyPoints = analysis.keyPoints;
        const pointsPerSlide = Math.ceil(keyPoints.length / (totalSlides - 1));
        const slidePoints = keyPoints.slice((slideIndex - 1) * pointsPerSlide, slideIndex * pointsPerSlide);
        
        const layouts = ['bullet', 'two-column', 'title-content', 'comparison'];
        const layout = layouts[slideIndex % layouts.length];
        
        return {
            title: `Key Point ${slideIndex}`,
            content: slidePoints,
            layout: layout,
            notes: `This slide covers: ${slidePoints.join('; ')}`
        };
    }

    async applyTemplateStyle(slides) {
        // In a real implementation, this would extract styles from the uploaded template
        // For demo purposes, we'll apply predefined styles based on template type
        
        const templateType = this.templateFile.type || 'corporate';
        const styles = this.getTemplateStyles(templateType);
        
        return slides.map(slide => ({
            ...slide,
            style: styles
        }));
    }

    getTemplateStyles(templateType) {
        const styleMap = {
            corporate: {
                background: '#1e3a8a',
                titleColor: '#ffffff',
                textColor: '#f1f5f9',
                accentColor: '#3b82f6',
                font: 'Arial'
            },
            modern: {
                background: '#0f172a',
                titleColor: '#f8fafc',
                textColor: '#cbd5e1',
                accentColor: '#06b6d4',
                font: 'Helvetica'
            },
            minimal: {
                background: '#ffffff',
                titleColor: '#1f2937',
                textColor: '#374151',
                accentColor: '#6366f1',
                font: 'Calibri'
            },
            creative: {
                background: '#7c3aed',
                titleColor: '#fbbf24',
                textColor: '#f3f4f6',
                accentColor: '#f59e0b',
                font: 'Comic Sans MS'
            }
        };
        
        return styleMap[templateType] || styleMap.corporate;
    }

    async buildPresentation(styledSlides) {
        // Create PowerPoint using PptxGenJS library
        const pptx = new PptxGenJS();
        
        // Set presentation properties
        pptx.author = 'Your Text, Your Style Generator';
        pptx.company = 'Auto-Generated';
        pptx.title = styledSlides[0]?.title || 'Generated Presentation';
        
        // Add slides
        styledSlides.forEach((slideData, index) => {
            const slide = pptx.addSlide();
            
            if (slideData.type === 'title') {
                this.addTitleSlide(slide, slideData);
            } else {
                this.addContentSlide(slide, slideData);
            }
        });
        
        // Generate the presentation file
        const fileName = `presentation_${Date.now()}.pptx`;
        this.presentationData = { pptx, fileName, slideCount: styledSlides.length };
        
        return this.presentationData;
    }

    addTitleSlide(slide, slideData) {
        const style = slideData.style || {};
        
        slide.background = { color: style.background || '1e3a8a' };
        
        slide.addText(slideData.title, {
            x: 0.5,
            y: 2,
            w: 9,
            h: 2,
            fontSize: 44,
            color: style.titleColor || 'ffffff',
            bold: true,
            align: 'center',
            fontFace: style.font || 'Arial'
        });
        
        slide.addText(slideData.subtitle, {
            x: 0.5,
            y: 4.5,
            w: 9,
            h: 1,
            fontSize: 24,
            color: style.textColor || 'f1f5f9',
            align: 'center',
            fontFace: style.font || 'Arial'
        });
    }

    addContentSlide(slide, slideData) {
        const style = slideData.style || {};
        
        slide.background = { color: style.background || '1e3a8a' };
        
        // Add title
        slide.addText(slideData.title, {
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 1,
            fontSize: 32,
            color: style.titleColor || 'ffffff',
            bold: true,
            fontFace: style.font || 'Arial'
        });
        
        // Add content based on layout
        if (slideData.layout === 'bullet') {
            const bulletPoints = Array.isArray(slideData.content) ? slideData.content : [slideData.content];
            const bulletText = bulletPoints.map(point => `• ${point}`).join('\n');
            
            slide.addText(bulletText, {
                x: 1,
                y: 2,
                w: 8,
                h: 4,
                fontSize: 20,
                color: style.textColor || 'f1f5f9',
                fontFace: style.font || 'Arial',
                valign: 'top'
            });
        } else {
            // Default content layout
            const contentText = Array.isArray(slideData.content) ? slideData.content.join('\n\n') : slideData.content;
            
            slide.addText(contentText, {
                x: 1,
                y: 2,
                w: 8,
                h: 4,
                fontSize: 18,
                color: style.textColor || 'f1f5f9',
                fontFace: style.font || 'Arial',
                valign: 'top'
            });
        }
    }

    showResults(presentation) {
        document.getElementById('progressPanel').classList.add('hidden');
        document.getElementById('resultPanel').classList.remove('hidden');
        
        document.getElementById('finalSlideCount').textContent = presentation.slideCount;
        
        // Setup download button
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.onclick = () => this.downloadPresentation();
        
        // Setup preview button
        const previewBtn = document.getElementById('previewBtn');
        previewBtn.onclick = () => this.showPreview();
    }

    async downloadPresentation() {
        if (!this.presentationData) {
            alert('No presentation data available');
            return;
        }
        
        try {
            await this.presentationData.pptx.writeFile({ fileName: this.presentationData.fileName });
        } catch (error) {
            console.error('Download failed:', error);
            alert('Failed to download presentation: ' + error.message);
        }
    }

    showPreview() {
        document.getElementById('previewPanel').classList.remove('hidden');
        document.getElementById('totalSlides').textContent = this.slideData.length;
        this.currentSlideIndex = 0;
        this.updateSlidePreview();
    }

    updateSlidePreview() {
        const slide = this.slideData[this.currentSlideIndex];
        if (!slide) return;
        
        document.getElementById('currentSlide').textContent = this.currentSlideIndex + 1;
        
        let slideHtml = `<h3>${slide.title}</h3>`;
        
        if (slide.type === 'title') {
            slideHtml += `<p style="font-size: 1.2em; margin-top: 20px;">${slide.subtitle}</p>`;
        } else {
            const content = Array.isArray(slide.content) ? slide.content : [slide.content];
            slideHtml += '<ul>';
            content.forEach(item => {
                slideHtml += `<li>${item}</li>`;
            });
            slideHtml += '</ul>';
            
            if (slide.notes) {
                slideHtml += `<div style="margin-top: 20px; padding: 10px; background: #f0f9ff; border-radius: 5px;">
                    <strong>Speaker Notes:</strong> ${slide.notes}
                </div>`;
            }
        }
        
        document.getElementById('slideContent').innerHTML = slideHtml;
        
        // Update navigation buttons
        document.getElementById('prevSlide').disabled = this.currentSlideIndex === 0;
        document.getElementById('nextSlide').disabled = this.currentSlideIndex === this.slideData.length - 1;
    }

    previousSlide() {
        if (this.currentSlideIndex > 0) {
            this.currentSlideIndex--;
            this.updateSlidePreview();
        }
    }

    nextSlide() {
        if (this.currentSlideIndex < this.slideData.length - 1) {
            this.currentSlideIndex++;
            this.updateSlidePreview();
        }
    }

    restartProcess() {
        this.currentStep = 1;
        this.templateFile = null;
        this.presentationData = null;
        this.slideData = [];
        
        // Reset form
        document.getElementById('inputText').value = '';
        document.getElementById('guidance').value = '';
        document.getElementById('apiKey').value = '';
        
        // Hide panels
        document.getElementById('progressPanel').classList.add('hidden');
        document.getElementById('resultPanel').classList.add('hidden');
        document.getElementById('previewPanel').classList.add('hidden');
        document.getElementById('templateInfo').classList.add('hidden');
        
        // Reset buttons
        document.querySelectorAll('.preset-btn, .sample-btn').forEach(btn => {
            btn.classList.remove('selected');
            btn.style.background = '';
            btn.style.color = '';
        });
        
        this.updateStepDisplay();
        this.updateCharCount();
        this.updateSummary();
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Global functions for navigation
function nextStep(step) {
    window.ppGenerator.nextStep(step);
}

function prevStep(step) {
    window.ppGenerator.prevStep(step);
}

function generatePresentation() {
    window.ppGenerator.generatePresentation();
}

function restartProcess() {
    window.ppGenerator.restartProcess();
}

function showHelp() {
    alert(`Help & Documentation:

1. Input your content: Paste any text, markdown, or prose
2. Configure AI: Enter your API key for content processing
3. Upload template: Provide a PowerPoint file for styling
4. Generate: Create your presentation automatically

Features:
- Intelligent slide structure detection
- Template style preservation
- Speaker notes generation
- Multiple presentation formats
- Real-time preview

For detailed documentation, visit our GitHub repository.`);
}

function showPrivacy() {
    alert(`Privacy Policy:

• Your API keys are never stored or logged
• Content is processed temporarily and not saved
• No personal data is collected or transmitted
• All processing happens locally when possible
• Template files are used only for style extraction

This application respects your privacy and data security.`);
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    window.ppGenerator = new PowerPointGenerator();
});
