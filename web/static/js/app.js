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
            selectedShapeIndex: null
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
        async loadPresentationForEditing() {
            if (!this.selectedJobForEditor) {
                this.editingPresentation = null;
                return;
            }

            console.log('Loading presentation for job:', this.selectedJobForEditor); // Debug

            try {
                const response = await fetch(`/api/jobs/${this.selectedJobForEditor}/presentation`);
                if (response.ok) {
                    this.editingPresentation = await response.json();
                    console.log('Loaded presentation data:', this.editingPresentation); // Debug
                    this.selectedShape = null;
                } else {
                    const error = await response.json();
                    console.error('API error:', error); // Debug
                    this.showAlert(`Error loading presentation: ${error.detail}`, 'danger');
                }
            } catch (error) {
                console.error('Error loading presentation:', error);
                this.showAlert(`Error loading presentation: ${error.message}`, 'danger');
            }
        },

        selectShape(slideIndex, shapeIndex) {
            this.selectedSlideIndex = slideIndex;
            this.selectedShapeIndex = shapeIndex;
            this.selectedShape = this.editingPresentation.slides[slideIndex].shapes[shapeIndex];
        },

        getShapeStyle(shape) {
            // Convert EMU (English Metric Units) to pixels for display
            // Use actual slide dimensions from the loaded presentation

            console.log('Shape data:', shape); // Debug log

            if (!this.editingPresentation || !this.editingPresentation.slides || !this.editingPresentation.slides[0]) {
                console.warn('No presentation data available for shape styling');
                return {};
            }

            const slideWidth = this.editingPresentation.slides[0].width;
            const slideHeight = this.editingPresentation.slides[0].height;
            const canvasWidth = 800; // Our display canvas width
            const canvasHeight = 600; // Our display canvas height

            const scaleX = canvasWidth / slideWidth;
            const scaleY = canvasHeight / slideHeight;

            const left = Math.round((shape.left || 0) * scaleX);
            const top = Math.round((shape.top || 0) * scaleY);
            const width = Math.max(50, Math.round((shape.width || 100000) * scaleX));
            const height = Math.max(20, Math.round((shape.height || 50000) * scaleY));

            const style = {
                left: `${left}px`,
                top: `${top}px`,
                width: `${width}px`,
                height: `${height}px`,
                backgroundColor: 'rgba(0, 123, 255, 0.1)',
                cursor: 'pointer',
                zIndex: '1'
            };

            console.log('Computed style:', style); // Debug log
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
        }
    }
}).mount('#app');