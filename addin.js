/**
 * Geotab DVIR Email Manager Add-in
 * @returns {{initialize: Function, focus: Function, blur: Function}}
 */
geotab.addin.dvirEmailer = function () {
    'use strict';

    let api;
    let state;
    let elAddin;
    let currentDatabase = null;
    let recipients = [];

    /**
     * Ensure current database is in Firestore
     */
    async function ensureDatabaseInFirestore() {
        if (!api || !window.db) {
            return;
        }
        
        try {
            api.getSession(async function(session) {
                currentDatabase = session.database;
                document.getElementById('databaseName').textContent = currentDatabase;
                
                if (currentDatabase && currentDatabase !== 'demo') {
                    // Check if database already exists
                    const querySnapshot = await window.db.collection('geotab_databases')
                        .where('database_name', '==', currentDatabase)
                        .get();
                    
                    if (querySnapshot.empty) {
                        // Add new database
                        await window.db.collection('geotab_databases').add({
                            database_name: currentDatabase,
                            added_at: firebase.firestore.FieldValue.serverTimestamp(),
                            active: true
                        });
                        console.log(`Added database ${currentDatabase} to Firestore`);
                    }
                }
            });
        } catch (error) {
            console.error('Error updating settings:', error);
            showAlert('Error updating settings: ' + error.message, 'danger');
        }
    }

    /**
     * Render recipients list
     */
    function renderRecipients() {
        const container = document.getElementById('recipientsList');
        
        if (recipients.length === 0) {
            showEmptyRecipientsState();
            return;
        }
        
        const recipientsHtml = recipients.map(recipient => `
            <div class="recipient-item">
                <div class="recipient-email">${recipient.email}</div>
                <button class="btn btn-outline-danger btn-sm" onclick="confirmRemoveRecipient('${recipient.id}', '${recipient.email}')">
                    <i class="fas fa-trash me-1"></i>Remove
                </button>
            </div>
        `).join('');
        
        container.innerHTML = recipientsHtml;
    }

    /**
     * Show empty recipients state
     */
    function showEmptyRecipientsState() {
        const container = document.getElementById('recipientsList');
        container.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-inbox"></i>
                <h5>No recipients configured</h5>
                <p>Add email addresses above to receive DVIR defect notifications</p>
            </div>
        `;
    }

    /**
     * Update recipient count display
     */
    function updateRecipientCount() {
        document.getElementById('recipientCount').textContent = recipients.length;
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
        
        // Auto-remove after 3 seconds
        setTimeout(() => {
            const alert = document.getElementById(alertId);
            if (alert && typeof bootstrap !== 'undefined' && bootstrap.Alert) {
                const bsAlert = new bootstrap.Alert(alert);
                bsAlert.close();
            }
        }, 3000);
    }

    /**
     * Setup event listeners
     */
    function setupEventListeners() {
        // Add recipient form
        const addRecipientForm = document.getElementById('addRecipientForm');
        if (addRecipientForm) {
            addRecipientForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const emailInput = document.getElementById('emailInput');
                const email = emailInput.value.trim();
                
                if (email) {
                    const success = await addRecipient(email);
                    if (success) {
                        emailInput.value = '';
                    }
                }
            });
        }
        
        // Settings switch
        const settingsSwitch = document.getElementById('onlyNewDefectsSwitch');
        if (settingsSwitch) {
            settingsSwitch.addEventListener('change', updateSettings);
        }
    }

    /**
     * Confirm recipient removal
     */
    window.confirmRemoveRecipient = function(recipientId, email) {
        if (confirm(`Are you sure you want to remove ${email} from the recipient list?`)) {
            removeRecipient(recipientId, email);
        }
    };

    /**
     * Refresh recipients
     */
    window.refreshRecipients = function() {
        loadRecipients();
    };

    /**
     * Test email system
     */
    window.testEmailSystem = function() {
        if (recipients.length === 0) {
            showAlert('No recipients configured. Add at least one recipient to test the email system.', 'warning');
            return;
        }
        
        // This would typically trigger a test email through your backend
        showAlert('Test email functionality would be implemented in your backend service', 'info');
    };

    /**
     * Export settings
     */
    window.exportSettings = function() {
        const settingsData = {
            database: currentDatabase,
            recipients: recipients.map(r => ({
                email: r.email,
                send_only_new_defects: r.send_only_new_defects
            })),
            settings: {
                send_only_new_defects: document.getElementById('onlyNewDefectsSwitch').checked
            },
            exported_at: new Date().toISOString()
        };
        
        const blob = new Blob([JSON.stringify(settingsData, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `dvir-email-settings-${currentDatabase}-${new Date().toISOString().split('T')[0]}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showAlert('Settings exported successfully', 'success');
    };

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
            }, 1000);
            
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

    /**
     * Load recipients from Firestore
     */
    async function loadRecipients() {
        if (!window.db || !currentDatabase) {
            return;
        }
        
        try {
            showAlert('Loading recipients...', 'info');
            
            const recipientsQuery = window.db.collection('dvir_recipients')
                .where('database_name', '==', currentDatabase);
            
            const snapshot = await recipientsQuery.get();
            recipients = [];
            let settings = { send_only_new_defects: true };
            
            snapshot.forEach((doc) => {
                const data = doc.to_dict();
                recipients.push({
                    id: doc.id,
                    email: data.email,
                    database_name: data.database_name,
                    send_only_new_defects: data.send_only_new_defects,
                    created_at: data.created_at
                });
                
                // Get settings from any recipient (should be consistent)
                if (data.send_only_new_defects !== undefined) {
                    settings.send_only_new_defects = data.send_only_new_defects;
                }
            });
            
            // Update UI
            document.getElementById('onlyNewDefectsSwitch').checked = settings.send_only_new_defects;
            renderRecipients();
            updateRecipientCount();
            showAlert(`Loaded ${recipients.length} recipients`, 'success');
            
        } catch (error) {
            console.error('Error loading recipients:', error);
            showAlert('Error loading recipients: ' + error.message, 'danger');
            showEmptyRecipientsState();
        }
    }

    /**
     * Add a new recipient
     */
    async function addRecipient(email) {
        if (!window.db || !currentDatabase) {
            showAlert('Database not initialized', 'danger');
            return false;
        }
        
        try {
            // Check if recipient already exists
            const existingQuery = window.db.collection('dvir_recipients')
                .where('database_name', '==', currentDatabase)
                .where('email', '==', email);
            
            const existingSnapshot = await existingQuery.get();
            if (!existingSnapshot.empty) {
                showAlert('This email address is already added as a recipient', 'warning');
                return false;
            }
            
            showAlert('Adding recipient...', 'info');
            
            // Add new recipient
            const recipientData = {
                email: email,
                database_name: currentDatabase,
                send_only_new_defects: document.getElementById('onlyNewDefectsSwitch').checked,
                created_at: firebase.firestore.FieldValue.serverTimestamp()
            };
            
            const docRef = await window.db.collection('dvir_recipients').add(recipientData);
            
            // Add to local array
            recipients.push({
                id: docRef.id,
                ...recipientData
            });
            
            renderRecipients();
            updateRecipientCount();
            showAlert(`Successfully added ${email} as a recipient`, 'success');
            return true;
            
        } catch (error) {
            console.error('Error adding recipient:', error);
            showAlert('Error adding recipient: ' + error.message, 'danger');
            return false;
        }
    }

    /**
     * Remove a recipient
     */
    async function removeRecipient(recipientId, email) {
        if (!window.db) {
            showAlert('Database not initialized', 'danger');
            return false;
        }
        
        try {
            showAlert('Removing recipient...', 'info');
            
            // Remove from Firestore
            await window.db.collection('dvir_recipients').doc(recipientId).delete();
            
            // Remove from local array
            recipients = recipients.filter(r => r.id !== recipientId);
            
            renderRecipients();
            updateRecipientCount();
            showAlert(`Successfully removed ${email}`, 'success');
            return true;
            
        } catch (error) {
            console.error('Error removing recipient:', error);
            showAlert('Error removing recipient: ' + error.message, 'danger');
            return false;
        }
    }

    /**
     * Update settings for all recipients in the database
     */
    async function updateSettings() {
        if (!window.db || !currentDatabase || recipients.length === 0) {
            return;
        }
        
        try {
            const sendOnlyNewDefects = document.getElementById('onlyNewDefectsSwitch').checked;
            showAlert('Updating settings...', 'info');
            
            // Update all recipients for this database
            const batch = window.db.batch();
            
            recipients.forEach(recipient => {
                const docRef = window.db.collection('dvir_recipients').doc(recipient.id);
                batch.update(docRef, {
                    send_only_new_defects: sendOnlyNewDefects
                });
            });
            
            await batch.commit();
            
            // Update local data
            recipients.forEach(recipient => {
                recipient.send_only_new_defects = sendOnlyNewDefects;
            });
            
            showAlert('Settings updated successfully', 'success');
            
        } catch (error) {
            console.error('Error ensuring database in Firestore:', error);
        }
    }