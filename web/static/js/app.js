// MCP PowerPoint Tools - Vue.js Application
const { createApp } = Vue;

createApp({
    data() {
        return {
            // File handling
            selectedPresentationFile: null,
            selectedDocumentFiles: [],

            // Processing state
            processing: false,
            ragProcessing: false,
            searching: false,

            // Processing options
            processingOptions: {
                includeAnalysis: false,
                namingStrategy: 'hybrid',
                useRag: true
            },

            // Current job tracking
            currentJob: null,
            pollInterval: null,
            websocket: null,

            // Jobs list
            jobs: [],

            // Search
            searchQuery: '',
            searchResults: [],
            searchPerformed: false,
            lastSearchQuery: '',

            // RAG system
            ragStats: null,

            // UI state
            alertMessage: '',
            alertType: 'info',

            // Shape Editor
            selectedJobForEditor: '',
            editingPresentation: null,
            selectedShape: null,
            selectedSlideIndex: null,
            selectedShapeIndex: null,
            presentationLoadingStatus: 'idle', // 'idle', 'loading', 'loaded', 'error', 'empty'
            slideCount: 0,

            // Document context system
            availableDocuments: [],
            selectedDocuments: [],
            isGeneratingContext: false,
            isGeneratingText: false
        };
    },

    computed: {
        progressPercent() {
            if (!this.currentJob) return 0;
            return Math.round((this.currentJob.steps_completed / this.currentJob.total_steps) * 100);
        },

        progressBarClass() {
            if (!this.currentJob) return '';
            if (this.currentJob.status === 'completed') return 'bg-success';
            if (this.currentJob.status === 'error') return 'bg-danger';
            return '';
        },

        completedJobs() {
            return this.jobs.filter(job => job.status === 'completed');
        },

        isLoadingPresentation() {
            return this.presentationLoadingStatus === 'loading';
        },

        hasValidPresentation() {
            return this.presentationLoadingStatus === 'loaded' && this.slideCount > 0;
        },

        hasEmptyPresentation() {
            return this.presentationLoadingStatus === 'empty' ||
                   (this.presentationLoadingStatus === 'loaded' && this.slideCount === 0);
        },

        hasLoadingError() {
            return this.presentationLoadingStatus === 'error';
        }
    },

    mounted() {
        this.loadJobs();
        this.loadRagStats();
        this.initWebSocket();
    },

    beforeUnmount() {
        if (this.pollInterval) {
            clearInterval(this.pollInterval);
        }
        if (this.websocket) {
            this.websocket.close();
        }
    },

    watch: {
        selectedJobForEditor(newVal, oldVal) {
            console.log('selectedJobForEditor changed:', { from: oldVal, to: newVal });
            // Only reset state when changing to a different job, not when loading completes
            if (oldVal !== newVal && oldVal !== '' && newVal !== '') {
                console.log('Resetting presentation state due to job change');
                this.resetPresentationState();
            }
        },
        presentationLoadingStatus(newVal, oldVal) {
            console.log('presentationLoadingStatus changed:', { from: oldVal, to: newVal });
        },
        slideCount(newVal, oldVal) {
            console.log('slideCount changed:', { from: oldVal, to: newVal });
        },
        hasValidPresentation(newVal) {
            console.log('hasValidPresentation:', newVal);
        },
        hasEmptyPresentation(newVal) {
            console.log('hasEmptyPresentation:', newVal);
        },
        isLoadingPresentation(newVal) {
            console.log('isLoadingPresentation:', newVal);
        }
    },

    methods: {
        // File handling
        handlePresentationSelect(event) {
            this.selectedPresentationFile = event.target.files[0];
        },

        handleDocumentSelect(event) {
            this.selectedDocumentFiles = Array.from(event.target.files);
        },

        // Main processing workflow
        async uploadAndProcess() {
            if (!this.selectedPresentationFile) return;

            this.processing = true;

            try {
                // Step 1: Upload presentation
                const uploadFormData = new FormData();
                uploadFormData.append('file', this.selectedPresentationFile);

                const uploadResponse = await fetch('/api/upload/presentation', {
                    method: 'POST',
                    body: uploadFormData
                });

                if (!uploadResponse.ok) {
                    const error = await uploadResponse.json();
                    throw new Error(error.detail || 'Upload failed');
                }

                const uploadResult = await uploadResponse.json();

                // Step 2: Start processing
                const processFormData = new FormData();
                processFormData.append('include_analysis', this.processingOptions.includeAnalysis);
                processFormData.append('naming_strategy', this.processingOptions.namingStrategy);
                processFormData.append('use_rag', this.processingOptions.useRag);

                const processResponse = await fetch(`/api/process/${uploadResult.job_id}`, {
                    method: 'POST',
                    body: processFormData
                });

                if (!processResponse.ok) {
                    const error = await processResponse.json();
                    throw new Error(error.detail || 'Processing failed');
                }

                // Step 3: Start polling for updates
                this.currentJob = uploadResult;
                this.startJobPolling(uploadResult.job_id);

                this.showAlert('Processing started successfully!', 'success');

            } catch (error) {
                console.error('Error:', error);
                this.showAlert(`Error: ${error.message}`, 'danger');
            } finally {
                this.processing = false;
            }
        },

        // Job polling (fallback when WebSocket is not available)
        startJobPolling(jobId) {
            // If WebSocket is connected, rely on real-time updates instead of polling
            if (this.websocket && this.websocket.readyState === WebSocket.OPEN) {
                console.log('Using WebSocket for real-time updates, skipping polling');
                return;
            }

            if (this.pollInterval) {
                clearInterval(this.pollInterval);
            }

            this.pollInterval = setInterval(async () => {
                try {
                    const response = await fetch(`/api/jobs/${jobId}/status`);
                    if (response.ok) {
                        const jobData = await response.json();
                        this.currentJob = jobData;

                        if (jobData.status === 'completed' || jobData.status === 'error') {
                            clearInterval(this.pollInterval);
                            this.pollInterval = null;
                            this.loadJobs(); // Refresh jobs list
                        }
                    }
                } catch (error) {
                    console.error('Error polling job status:', error);
                }
            }, 2000); // Poll every 2 seconds
        },

        // Document upload
        async uploadDocuments() {
            if (this.selectedDocumentFiles.length === 0) return;

            this.ragProcessing = true;

            try {
                const formData = new FormData();
                this.selectedDocumentFiles.forEach(file => {
                    formData.append('files', file);
                });

                const response = await fetch('/api/upload/documents', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.detail || 'Document upload failed');
                }

                const result = await response.json();

                this.showAlert(
                    `Uploaded ${result.uploaded_count} documents successfully!`,
                    'success'
                );

                // Clear file input
                document.getElementById('documentFiles').value = '';
                this.selectedDocumentFiles = [];

            } catch (error) {
                console.error('Error uploading documents:', error);
                this.showAlert(`Error uploading documents: ${error.message}`, 'danger');
            } finally {
                this.ragProcessing = false;
            }
        },

        // RAG operations
        async ingestDocuments() {
            this.ragProcessing = true;

            try {
                const response = await fetch('/api/rag/ingest', {
                    method: 'POST'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.detail || 'Document ingestion failed');
                }

                const result = await response.json();

                this.showAlert(result.message, 'success');
                this.loadRagStats();

            } catch (error) {
                console.error('Error ingesting documents:', error);
                this.showAlert(`Error ingesting documents: ${error.message}`, 'danger');
            } finally {
                this.ragProcessing = false;
            }
        },

        async loadRagStats() {
            try {
                const response = await fetch('/api/rag/stats');
                if (response.ok) {
                    this.ragStats = await response.json();
                }
            } catch (error) {
                console.error('Error loading RAG stats:', error);
            }
        },

        // Search functionality
        async performSearch() {
            if (!this.searchQuery.trim()) return;

            this.searching = true;
            this.lastSearchQuery = this.searchQuery;

            try {
                const response = await fetch(
                    `/api/rag/search?query=${encodeURIComponent(this.searchQuery)}&max_results=10`
                );

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.detail || 'Search failed');
                }

                const result = await response.json();
                this.searchResults = result.results || [];
                this.searchPerformed = true;

            } catch (error) {
                console.error('Error searching:', error);
                this.showAlert(`Search error: ${error.message}`, 'danger');
                this.searchResults = [];
            } finally {
                this.searching = false;
            }
        },

        // Jobs management
        async loadJobs() {
            try {
                const response = await fetch('/api/jobs');
                if (response.ok) {
                    const data = await response.json();
                    this.jobs = data.jobs || [];
                }
            } catch (error) {
                console.error('Error loading jobs:', error);
            }
        },

        async deleteJob(jobId) {
            if (!confirm('Are you sure you want to delete this job?')) return;

            try {
                const response = await fetch(`/api/jobs/${jobId}`, {
                    method: 'DELETE'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.detail || 'Delete failed');
                }

                this.showAlert('Job deleted successfully!', 'success');
                this.loadJobs();

                // Clear current job if it was deleted
                if (this.currentJob && this.currentJob.job_id === jobId) {
                    this.currentJob = null;
                    if (this.pollInterval) {
                        clearInterval(this.pollInterval);
                        this.pollInterval = null;
                    }
                }

            } catch (error) {
                console.error('Error deleting job:', error);
                this.showAlert(`Error deleting job: ${error.message}`, 'danger');
            }
        },

        downloadFile(jobId, fileType) {
            const url = `/api/download/${jobId}/${fileType}`;
            const link = document.createElement('a');
            link.href = url;
            link.download = '';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        },

        // Utility functions
        formatFileSize(bytes) {
            if (!bytes) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
        },

        formatDate(dateString) {
            if (!dateString) return '';
            return new Date(dateString).toLocaleString();
        },

        getJobProgress(job) {
            if (!job.total_steps) return 0;
            return Math.round((job.steps_completed / job.total_steps) * 100);
        },

        getStatusColor(status) {
            switch (status) {
                case 'completed': return 'success';
                case 'processing': return 'warning';
                case 'error': return 'danger';
                case 'uploaded': return 'info';
                default: return 'secondary';
            }
        },

        // Alert management
        showAlert(message, type = 'info') {
            this.alertMessage = message;
            this.alertType = type;

            // Auto-clear success and info alerts
            if (type === 'success' || type === 'info') {
                setTimeout(() => {
                    this.clearAlert();
                }, 5000);
            }
        },

        clearAlert() {
            this.alertMessage = '';
            this.alertType = 'info';
        },

        // WebSocket functionality
        initWebSocket() {
            const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
            const host = window.location.host;
            this.websocket = new WebSocket(`${protocol}//${host}/ws`);

            this.websocket.onopen = () => {
                console.log('WebSocket connected');
                this.showAlert('Real-time updates connected', 'success');
            };

            this.websocket.onmessage = (event) => {
                const message = JSON.parse(event.data);
                this.handleWebSocketMessage(message);
            };

            this.websocket.onclose = () => {
                console.log('WebSocket disconnected');
                // Attempt to reconnect after 3 seconds
                setTimeout(() => {
                    this.initWebSocket();
                }, 3000);
            };

            this.websocket.onerror = (error) => {
                console.error('WebSocket error:', error);
            };
        },

        handleWebSocketMessage(message) {
            if (message.type === 'job_update') {
                const jobData = message.data;

                // Update current job if it matches
                if (this.currentJob && this.currentJob.job_id === jobData.job_id) {
                    this.currentJob = jobData;
                    this.animateProgressUpdate();

                    // Show completion notification
                    if (jobData.status === 'completed') {
                        this.showAlert(`Processing completed for ${jobData.filename}!`, 'success');
                        // Clear polling interval since we have real-time updates
                        if (this.pollInterval) {
                            clearInterval(this.pollInterval);
                            this.pollInterval = null;
                        }
                    } else if (jobData.status === 'error') {
                        this.showAlert(`Processing failed for ${jobData.filename}: ${jobData.error}`, 'danger');
                        if (this.pollInterval) {
                            clearInterval(this.pollInterval);
                            this.pollInterval = null;
                        }
                    }
                }

                // Update jobs list
                this.updateJobInList(jobData);
            }
        },

        updateJobInList(jobData) {
            const index = this.jobs.findIndex(job => job.job_id === jobData.job_id);
            if (index >= 0) {
                this.jobs[index] = jobData;
            } else {
                this.jobs.unshift(jobData);
            }
        },

        animateProgressUpdate() {
            // Add a CSS class to trigger progress bar animation
            this.$nextTick(() => {
                const progressBar = document.querySelector('.progress-bar');
                if (progressBar) {
                    progressBar.classList.add('updating');
                    setTimeout(() => {
                        progressBar.classList.remove('updating');
                    }, 1000);
                }
            });
        },

        // Shape Editor functionality
        countSlidesInPresentation(presentationData) {
            if (!presentationData || !Array.isArray(presentationData.slides)) {
                return 0;
            }
            return presentationData.slides.length;
        },

        resetPresentationState() {
            this.editingPresentation = null;
            this.selectedShape = null;
            this.selectedSlideIndex = null;
            this.selectedShapeIndex = null;
            this.slideCount = 0;
            this.presentationLoadingStatus = 'idle';
        },

        async loadPresentationForEditing() {
            if (!this.selectedJobForEditor) {
                this.resetPresentationState();
                return;
            }

            // Prevent multiple concurrent loads
            if (this.presentationLoadingStatus === 'loading') {
                console.log('Already loading presentation, skipping...');
                return;
            }

            console.log('Loading presentation for job:', this.selectedJobForEditor);

            // Set loading state
            this.presentationLoadingStatus = 'loading';
            this.editingPresentation = null;
            this.selectedShape = null;
            this.slideCount = 0;

            try {
                const response = await fetch(`/api/jobs/${this.selectedJobForEditor}/presentation`);
                if (response.ok) {
                    const presentationData = await response.json();
                    console.log('Raw presentation data:', presentationData);

                    // Count slides first
                    const slideCount = this.countSlidesInPresentation(presentationData);
                    console.log('Slide count:', slideCount);

                    // Validate and normalize the presentation data
                    if (this.validatePresentationData(presentationData)) {
                        this.editingPresentation = this.normalizePresentationData(presentationData);
                        this.slideCount = this.countSlidesInPresentation(this.editingPresentation);

                        console.log('Normalized presentation data:', this.editingPresentation);
                        console.log('Final slide count:', this.slideCount);

                        // Set appropriate status based on slide count
                        if (this.slideCount > 0) {
                            this.presentationLoadingStatus = 'loaded';
                            console.log(`Successfully loaded presentation with ${this.slideCount} slides`);
                            this.showAlert(`Presentation loaded: ${this.slideCount} slide(s)`, 'success');
                            // Load available documents for context generation
                            this.loadAvailableDocuments();
                        } else {
                            this.presentationLoadingStatus = 'empty';
                            console.log('Presentation loaded but contains no slides');
                            this.showAlert('Presentation contains no slides', 'warning');
                        }
                    } else {
                        console.error('Presentation data validation failed');
                        this.presentationLoadingStatus = 'error';
                        this.editingPresentation = null;
                        this.slideCount = 0;
                        this.showAlert('Invalid presentation data structure', 'danger');
                    }
                } else {
                    const error = await response.json();
                    console.error('API error:', error);
                    this.presentationLoadingStatus = 'error';
                    this.editingPresentation = null;
                    this.slideCount = 0;
                    this.showAlert(`Error loading presentation: ${error.detail}`, 'danger');
                }
            } catch (error) {
                console.error('Error loading presentation:', error);
                this.presentationLoadingStatus = 'error';
                this.editingPresentation = null;
                this.slideCount = 0;
                this.showAlert(`Error loading presentation: ${error.message}`, 'danger');
            }
        },

        validatePresentationData(data) {
            // Check if we have required structure
            if (!data || typeof data !== 'object') {
                console.error('Presentation data is not an object');
                return false;
            }

            // Check for error in the data
            if (data.error) {
                console.error('Presentation data contains error:', data.error);
                return false;
            }

            // Check for slides array
            if (!Array.isArray(data.slides)) {
                console.error('Presentation data missing slides array');
                return false;
            }

            console.log('Presentation data validation passed');
            return true;
        },

        normalizePresentationData(data) {
            // Ensure we have proper slide dimensions at the root level
            const normalizedData = {
                slide_width: data.slide_width || 9144000, // Default PowerPoint width in EMU
                slide_height: data.slide_height || 6858000, // Default PowerPoint height in EMU
                slides: []
            };

            // Normalize each slide
            data.slides.forEach((slide, index) => {
                const normalizedSlide = {
                    slide_number: slide.slide_number || (index + 1),
                    name: slide.name || null,
                    layout: slide.layout || null,
                    width: slide.width || normalizedData.slide_width,
                    height: slide.height || normalizedData.slide_height,
                    shapes: []
                };

                // Normalize shapes
                if (Array.isArray(slide.shapes)) {
                    slide.shapes.forEach((shape, shapeIndex) => {
                        const normalizedShape = {
                            shape_id: shape.shape_id || shape.id || shapeIndex,
                            name: shape.name || `Shape ${shapeIndex + 1}`,
                            descriptive_name: shape.descriptive_name || shape.name || `Shape ${shapeIndex + 1}`,
                            shape_type: shape.shape_type || 'UNKNOWN',
                            left: parseInt(shape.left) || 0,
                            top: parseInt(shape.top) || 0,
                            width: parseInt(shape.width) || 914400, // Default 1 inch in EMU
                            height: parseInt(shape.height) || 914400, // Default 1 inch in EMU
                            rotation: shape.rotation || 0,
                            text_frame: shape.text_frame || null,
                            table: shape.table || null,
                            image: shape.image || null,
                            confidence_score: shape.confidence_score || null,
                            context_analysis: shape.context_analysis || null,
                            semantic_tags: shape.semantic_tags || []
                        };

                        normalizedSlide.shapes.push(normalizedShape);
                    });
                }

                normalizedData.slides.push(normalizedSlide);
            });

            console.log('Data normalization completed');
            return normalizedData;
        },

        selectShape(slideIndex, shapeIndex) {
            this.selectedSlideIndex = slideIndex;
            this.selectedShapeIndex = shapeIndex;
            this.selectedShape = this.editingPresentation.slides[slideIndex].shapes[shapeIndex];
        },

        getShapeStyle(shape) {
            // Convert EMU (English Metric Units) to pixels for display
            // Use actual slide dimensions from the loaded presentation

            console.log('Shape data for styling:', shape);

            if (!this.editingPresentation) {
                console.warn('No presentation data available for shape styling');
                return { display: 'none' };
            }

            if (!shape) {
                console.warn('No shape data provided for styling');
                return { display: 'none' };
            }

            // Get slide dimensions - check multiple possible locations
            let slideWidth = this.editingPresentation.slide_width;
            let slideHeight = this.editingPresentation.slide_height;

            // Fallback to first slide dimensions if not at root level
            if ((!slideWidth || !slideHeight) && this.editingPresentation.slides && this.editingPresentation.slides[0]) {
                slideWidth = this.editingPresentation.slides[0].width;
                slideHeight = this.editingPresentation.slides[0].height;
            }

            // Final fallback to default PowerPoint dimensions in EMU
            if (!slideWidth || !slideHeight || slideWidth === 0 || slideHeight === 0) {
                slideWidth = 9144000;  // 10 inches in EMU
                slideHeight = 6858000; // 7.5 inches in EMU
                console.warn('Using default slide dimensions:', { slideWidth, slideHeight });
            }

            // Get actual canvas size (responsive)
            const canvasElement = document.querySelector('.slide-canvas');
            const canvasWidth = canvasElement ? canvasElement.offsetWidth : 600;
            const canvasHeight = canvasElement ? canvasElement.offsetHeight : 450;

            // Calculate scale factors
            const scaleX = canvasWidth / slideWidth;
            const scaleY = canvasHeight / slideHeight;

            // Extract shape dimensions with defaults
            const shapeLeft = parseInt(shape.left) || 0;
            const shapeTop = parseInt(shape.top) || 0;
            const shapeWidth = parseInt(shape.width) || 914400; // Default 1 inch
            const shapeHeight = parseInt(shape.height) || 914400; // Default 1 inch

            // Convert shape dimensions from EMU to pixels
            const left = Math.round(shapeLeft * scaleX);
            const top = Math.round(shapeTop * scaleY);
            const width = Math.max(20, Math.round(shapeWidth * scaleX));
            const height = Math.max(20, Math.round(shapeHeight * scaleY));

            // Ensure shape is within canvas bounds
            const clampedLeft = Math.max(0, Math.min(left, canvasWidth - 20));
            const clampedTop = Math.max(0, Math.min(top, canvasHeight - 20));
            const clampedWidth = Math.min(width, canvasWidth - clampedLeft);
            const clampedHeight = Math.min(height, canvasHeight - clampedTop);

            const style = {
                left: `${clampedLeft}px`,
                top: `${clampedTop}px`,
                width: `${clampedWidth}px`,
                height: `${clampedHeight}px`,
                backgroundColor: 'rgba(0, 123, 255, 0.1)',
                border: '2px solid #007bff',
                cursor: 'pointer',
                zIndex: '1',
                overflow: 'hidden',
                fontSize: '11px'
            };

            console.log('Styling calculation:', {
                slideSize: { slideWidth, slideHeight },
                shapeEMU: { left: shapeLeft, top: shapeTop, width: shapeWidth, height: shapeHeight },
                scale: { scaleX, scaleY },
                pixelStyle: style
            });

            return style;
        },

        async saveShapeChanges() {
            if (!this.selectedShape || !this.editingPresentation) return;

            try {
                const response = await fetch(`/api/jobs/${this.selectedJobForEditor}/presentation`, {
                    method: 'PUT',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(this.editingPresentation)
                });

                if (response.ok) {
                    this.showAlert('Shape changes saved successfully!', 'success');
                } else {
                    const error = await response.json();
                    this.showAlert(`Error saving changes: ${error.detail}`, 'danger');
                }
            } catch (error) {
                console.error('Error saving shape changes:', error);
                this.showAlert(`Error saving changes: ${error.message}`, 'danger');
            }
        },

        // Context generation methods
        async generateContextFromName() {
            if (!this.selectedShape?.descriptive_name) {
                this.showAlert('Please enter a descriptive name first.', 'warning');
                return;
            }

            this.isGeneratingContext = true;
            try {
                const response = await fetch('/api/mcp/generate_context', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        descriptive_name: this.selectedShape.descriptive_name,
                        shape_type: this.selectedShape.shape_type || 'shape'
                    })
                });

                if (response.ok) {
                    const result = await response.json();
                    this.selectedShape.context = result.context;
                    this.showAlert('Context generated successfully!', 'success');
                } else {
                    const error = await response.json();
                    this.showAlert(`Error generating context: ${error.detail}`, 'danger');
                }
            } catch (error) {
                console.error('Error generating context:', error);
                this.showAlert(`Error generating context: ${error.message}`, 'danger');
            } finally {
                this.isGeneratingContext = false;
            }
        },

        async generateTextFromContext() {
            if (!this.selectedShape?.context) {
                this.showAlert('Please generate or enter context information first.', 'warning');
                return;
            }

            // Show warning if no documents selected but still proceed
            if (this.selectedDocuments.length === 0) {
                this.showAlert('No documents selected - generating basic content from context only. Select documents above for enhanced content.', 'info');
            }

            this.isGeneratingText = true;
            try {
                const response = await fetch('/api/mcp/generate_text_content', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        context: this.selectedShape.context,
                        selected_documents: this.selectedDocuments,
                        descriptive_name: this.selectedShape.descriptive_name || '',
                        shape_type: this.selectedShape.shape_type || 'shape'
                    })
                });

                if (response.ok) {
                    const result = await response.json();
                    // Ensure text_frame exists
                    if (!this.selectedShape.text_frame) {
                        this.selectedShape.text_frame = { text: '' };
                    }
                    this.selectedShape.text_frame.text = result.text_content;
                    this.showAlert('Text content generated successfully!', 'success');
                } else {
                    const error = await response.json();
                    this.showAlert(`Error generating text content: ${error.detail}`, 'danger');
                }
            } catch (error) {
                console.error('Error generating text content:', error);
                this.showAlert(`Error generating text content: ${error.message}`, 'danger');
            } finally {
                this.isGeneratingText = false;
            }
        },

        // Load available documents for context generation
        async loadAvailableDocuments() {
            try {
                const response = await fetch('/api/documents');
                if (response.ok) {
                    const documents = await response.json();
                    this.availableDocuments = documents.map(doc => ({
                        filename: doc.filename,
                        content_preview: doc.content ? doc.content.substring(0, 100) + '...' : 'No preview available'
                    }));
                } else {
                    console.log('No documents endpoint available or no documents found');
                    this.availableDocuments = [];
                }
            } catch (error) {
                console.error('Error loading documents:', error);
                this.availableDocuments = [];
            }
        },

        // Ensure text frame exists when editing text content
        ensureTextFrame() {
            if (this.selectedShape && !this.selectedShape.text_frame) {
                this.selectedShape.text_frame = { text: '' };
            }
        }
    }
}).mount('#app');