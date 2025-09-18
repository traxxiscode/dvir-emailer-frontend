/**
 * Geotab DVIR Emailer Add-in
 * @returns {{initialize: Function, focus: Function, blur: Function}}
 */
geotab.addin.dvirEmailer = function () {
    'use strict';

    let api;
    let state;
    let elAddin;
    let currentDatabase = null;

    /**
     * Make a Geotab API call
     */
    async function makeGeotabCall(method, typeName, parameters = {}) {
        if (!api) {
            throw new Error('Geotab API not initialized');
        }
        
        return new Promise((resolve, reject) => {
            const callParams = {
                typeName: typeName,
                ...parameters
            };
            
            api.call(method, callParams, resolve, reject);
        });
    }

    /**
     * Add current database to Firestore if it doesn't exist
     */
    async function ensureDatabaseInFirestore() {
        if (!api || !window.db) {
            return;
        }
        
        try {
            api.getSession(async function(session) {
                const databaseName = session.database;
                currentDatabase = databaseName;
                
                // Update UI with current database
                const dbElement = document.getElementById('currentDatabase');
                if (dbElement) {
                    dbElement.textContent = databaseName;
                }
                
                if (databaseName && databaseName !== 'demo') {
                    // Check if database configuration already exists
                    const querySnapshot = await window.db.collection('dvir_configurations')
                        .where('database_name', '==', databaseName)
                        .get();
                    
                    if (querySnapshot.empty) {
                        // Add new database configuration
                        await window.db.collection('dvir_configurations').add({
                            database_name: databaseName,
                            recipients: [],
                            created_at: firebase.firestore.FieldValue.serverTimestamp(),
                            updated_at: firebase.firestore.FieldValue.serverTimestamp(),
                            active: true
                        });
                        console.log(`Added database ${databaseName} configuration to Firestore`);
                    } else {
                        console.log(`Database ${databaseName} configuration already exists in Firestore`);
                    }
                }
            });
        } catch (error) {
            console.error('Error ensuring database in Firestore:', error);
            showAlert('Error connecting to database: ' + error.message, 'danger');
        }
    }

    /**
     * Load recipients for current database
     */
    async function loadRecipients() {
        if (!currentDatabase || !window.db) {
            showAlert('Database not initialized', 'danger');
            return;
        }

        try {
            const querySnapshot = await window.db.collection('dvir_configurations')
                .where('database_name', '==', currentDatabase)
                .get();
            
            if (!querySnapshot.empty) {
                const doc = querySnapshot.docs[0];
                const data = doc.data();
                const recipients = data.recipients || [];
                
                renderRecipients(recipients);
                updateRecipientCount(recipients.length);
            } else {
                renderRecipients([]);
                updateRecipientCount(0);
            }
        } catch (error) {
            console.error('Error loading recipients:', error);
            showAlert('Error loading recipients: ' + error.message, 'danger');
        }
    }

    /**
     * Add recipient to database
     */
    async function addRecipient(email, defectFilter) {
        if (!currentDatabase || !window.db) {
            showAlert('Database not initialized', 'danger');
            return;
        }

        try {
            const querySnapshot = await window.db.collection('dvir_configurations')
                .where('database_name', '==', currentDatabase)
                .get();
            
            if (!querySnapshot.empty) {
                const doc = querySnapshot.docs[0];
                const data = doc.data();
                const recipients = data.recipients || [];
                
                // Check if recipient already exists
                const existingRecipient = recipients.find(r => r.email === email);
                if (existingRecipient) {
                    showAlert('Recipient already exists', 'warning');
                    return;
                }
                
                // Add new recipient
                const newRecipient = {
                    email: email,
                    defect_filter: defectFilter,
                    added_at: firebase.firestore.FieldValue.serverTimestamp()
                };
                
                recipients.push(newRecipient);
                
                // Update document
                await doc.ref.update({
                    recipients: recipients,
                    updated_at: firebase.firestore.FieldValue.serverTimestamp()
                });
                
                showAlert(`Successfully added ${email} to recipient list`, 'success');
                loadRecipients(); // Refresh the list
                
                // Clear form
                document.getElementById('addRecipientForm').reset();
                
            } else {
                showAlert('Database configuration not found', 'danger');
            }
        } catch (error) {
            console.error('Error adding recipient:', error);
            showAlert('Error adding recipient: ' + error.message, 'danger');
        }
    }

    /**
     * Remove recipient from database
     */
    async function removeRecipient(email) {
        if (!currentDatabase || !window.db) {
            showAlert('Database not initialized', 'danger');
            return;
        }

        try {
            const querySnapshot = await window.db.collection('dvir_configurations')
                .where('database_name', '==', currentDatabase)
                .get();
            
            if (!querySnapshot.empty) {
                const doc = querySnapshot.docs[0];
                const data = doc.data();
                const recipients = data.recipients || [];
                
                // Remove recipient
                const updatedRecipients = recipients.filter(r => r.email !== email);
                
                // Update document
                await doc.ref.update({
                    recipients: updatedRecipients,
                    updated_at: firebase.firestore.FieldValue.serverTimestamp()
                });
                
                showAlert(`Successfully removed ${email} from recipient list`, 'success');
                loadRecipients(); // Refresh the list
                
            } else {
                showAlert('Database configuration not found', 'danger');
            }
        } catch (error) {
            console.error('Error removing recipient:', error);
            showAlert('Error removing recipient: ' + error.message, 'danger');
        }
    }

    /**
     * Render recipients list
     */
    function renderRecipients(recipients) {
        const container = document.getElementById('recipientsList');
        if (!container) return;
        
        if (recipients.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <i class="fas fa-inbox"></i>
                    <p>No recipients configured</p>
                    <small>Add email addresses to start receiving DVIR notifications</small>
                </div>
            `;
            return;
        }
        
        const recipientsHtml = recipients.map(recipient => `
            <div class="recipient-item">
                <div>
                    <div class="recipient-email">${recipient.email}</div>
                    <div class="recipient-settings">
                        <i class="fas fa-filter me-1"></i>
                        ${recipient.defect_filter === 'new' ? 'New Defects Only' : 'All Defects'}
                    </div>
                </div>
                <button class="btn btn-outline-danger btn-sm" onclick="confirmRemoveRecipient('${recipient.email}')">
                    <i class="fas fa-trash"></i> Remove
                </button>
            </div>
        `).join('');
        
        container.innerHTML = recipientsHtml;
    }

    /**
     * Update recipient count badge
     */
    function updateRecipientCount(count) {
        const countElement = document.getElementById('recipientCount');
        if (countElement) {
            countElement.textContent = count;
        }
    }

    /**
     * Show alert messages
     */
    function showAlert(message, type = 'info') {
        const alertContainer = document.getElementById('alertContainer');
        if (!alertContainer) return;
        
        const alertId = 'alert-' + Date.now();
        
        const iconMap = {
            'success': 'check-circle',
            'danger': 'exclamation-triangle',
            'warning': 'exclamation-triangle',
            'info': 'info-circle'
        };
        
        const alertHtml = `
            <div class="alert alert-${type} alert-dismissible fade show" id="${alertId}" role="alert">
                <i class="fas fa-${iconMap[type]} me-2"></i>
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;
        
        alertContainer.insertAdjacentHTML('beforeend', alertHtml);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            const alert = document.getElementById(alertId);
            if (alert && typeof bootstrap !== 'undefined' && bootstrap.Alert) {
                const bsAlert = new bootstrap.Alert(alert);
                bsAlert.close();
            }
        }, 5000);
    }

    /**
     * Test connection to Firestore
     */
    window.testConnection = async function() {
        try {
            showAlert('Testing connection...', 'info');
            
            if (!window.db) {
                showAlert('Firestore not initialized', 'danger');
                return;
            }
            
            // Try to read from the collection
            const snapshot = await window.db.collection('dvir_configurations').limit(1).get();
            showAlert('Connection test successful', 'success');
            
        } catch (error) {
            console.error('Connection test failed:', error);
            showAlert('Connection test failed: ' + error.message, 'danger');
        }
    };

    /**
     * Confirm recipient removal
     */
    window.confirmRemoveRecipient = function(email) {
        if (confirm(`Are you sure you want to remove ${email} from the recipient list?`)) {
            removeRecipient(email);
        }
    };

    /**
     * Refresh recipients list
     */
    window.refreshRecipients = function() {
        loadRecipients();
    };

    /**
     * Setup event listeners
     */
    function setupEventListeners() {
        // Add recipient form submission
        const addRecipientForm = document.getElementById('addRecipientForm');
        if (addRecipientForm) {
            addRecipientForm.addEventListener('submit', function(e) {
                e.preventDefault();
                
                const email = document.getElementById('recipientEmail').value.trim();
                const defectFilter = document.querySelector('input[name="defectFilter"]:checked').value;
                
                if (!email) {
                    showAlert('Please enter a valid email address', 'warning');
                    return;
                }
                
                addRecipient(email, defectFilter);
            });
        }
    }

    return {
        /**
         * initialize() is called only once when the Add-In is first loaded.
         */
        initialize: function (freshApi, freshState, initializeCallback) {
            api = freshApi;
            state = freshState;

            elAddin = document.getElementById('dvirEmailer');

            if (state.translate) {
                state.translate(elAddin || '');
            }
            
            initializeCallback();
        },

        /**
         * focus() is called whenever the Add-In receives focus.
         */
        focus: function (freshApi, freshState) {
            api = freshApi;
            state = freshState;

            // Ensure current database is in Firestore
            ensureDatabaseInFirestore();

            // Setup event listeners
            setupEventListeners();
            
            // Load recipients data
            setTimeout(() => {
                loadRecipients();
            }, 1000); // Give time for database to be set
            
            // Show main content
            if (elAddin) {
                elAddin.style.display = 'block';
            }
        },

        /**
         * blur() is called whenever the user navigates away from the Add-In.
         */
        blur: function () {
            // Hide main content
            if (elAddin) {
                elAddin.style.display = 'none';
            }
        }
    };
};