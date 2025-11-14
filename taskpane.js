// Configuration
const API_BASE_URL = 'https://bona-ulrike-uncoincided.ngrok-free.dev'; // Change this to your Flask API URL

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Outlook Add-in loaded successfully');
        
        // Set up event listeners
        document.getElementById('analyzeBtn').addEventListener('click', analyzeEmail);
    }
});

// Main analysis function
async function analyzeEmail() {
    const analyzeBtn = document.getElementById('analyzeBtn');
    const loading = document.getElementById('loading');
    const results = document.getElementById('results');
    const error = document.getElementById('error');
    const useClaude = document.getElementById('modelToggle').checked;

    // Reset UI
    error.classList.remove('active');
    results.classList.remove('active');
    
    // Show loading
    analyzeBtn.disabled = true;
    loading.classList.add('active');

    try {
        // Get the current email item
        const item = Office.context.mailbox.item;
        
        // Extract email data
        const emailData = await extractEmailData(item);
        
        // Call the Flask API
        const analysisResult = await callAnalysisAPI(emailData, useClaude);
        
        // Display results
        displayResults(analysisResult);
        
    } catch (err) {
        console.error('Analysis error:', err);
        showError(err.message || 'Failed to analyze email. Please try again.');
    } finally {
        analyzeBtn.disabled = false;
        loading.classList.remove('active');
    }
}

// Extract email data from Outlook
function extractEmailData(item) {
    return new Promise((resolve, reject) => {
        // Get subject
        const subject = item.subject || '';
        
        // Get sender
        const sender = item.from ? 
            `${item.from.displayName} <${item.from.emailAddress}>` : 
            item.sender ? `${item.sender.displayName} <${item.sender.emailAddress}>` : '';
        
        // Get body
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const body = result.value;
                
                // Get HTML body for link extraction
                item.body.getAsync(Office.CoercionType.Html, (htmlResult) => {
                    const htmlBody = htmlResult.status === Office.AsyncResultStatus.Succeeded ? 
                        htmlResult.value : '';
                    
                    resolve({
                        subject: subject,
                        sender: sender,
                        body: body,
                        htmlBody: htmlBody
                    });
                });
            } else {
                reject(new Error('Failed to read email body'));
            }
        });
    });
}

// Call the Flask API
async function callAnalysisAPI(emailData, useClaude) {
    try {
        const formData = new FormData();
        formData.append('use_claude', useClaude);
        
        // Create a pasted content string from the email data
        const pastedContent = `From: ${emailData.sender}\nSubject: ${emailData.subject}\n\n${emailData.body}`;
        formData.append('pasted_content', pastedContent);
        
        const response = await fetch(`${API_BASE_URL}/analyze`, {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Analysis failed');
        }

        const data = await response.json();
        return data;
        
    } catch (err) {
        if (err.message.includes('Failed to fetch')) {
            throw new Error('Cannot connect to analysis server. Make sure your Flask app is running.');
        }
        throw err;
    }
}

// Display analysis results
function displayResults(data) {
    const results = document.getElementById('results');
    const scoreValue = document.getElementById('scoreValue');
    const scoreBar = document.getElementById('scoreBar');
    const scoreLabel = document.getElementById('scoreLabel');
    const analysisText = document.getElementById('analysisText');
    const redFlags = document.getElementById('redFlags');
    const emailDetails = document.getElementById('emailDetails');

    // Show results section
    results.classList.add('active');

    // Confidence score
    const score = data.confidence_score;
    scoreValue.textContent = `${score}%`;
    scoreBar.style.width = `${score}%`;

    // Score label and badge
    let label = '';
    let badgeClass = '';
    if (score < 30) {
        label = '✅ Low Risk - Appears legitimate';
        badgeClass = 'badge-success';
    } else if (score < 60) {
        label = '⚠️ Medium Risk - Exercise caution';
        badgeClass = 'badge-warning';
    } else {
        label = '🚨 High Risk - Likely phishing';
        badgeClass = 'badge-danger';
    }
    scoreLabel.innerHTML = `<span class="badge ${badgeClass}">${label}</span>`;

    // Analysis text - take first 2-3 paragraphs or key points
    const analysisParagraphs = data.analysis.split('\n\n').slice(0, 2);
    analysisText.innerHTML = analysisParagraphs
        .map(p => `<p>${escapeHtml(p)}</p>`)
        .join('');

    // Build red flags list
    const flags = [];
    
    // Add sender issues
    if (data.sender_analysis && data.sender_analysis.length > 0) {
        data.sender_analysis.forEach(issue => {
            if (!issue.includes('No obvious') && !issue.includes('No sender issues')) {
                flags.push(`👤 ${issue}`);
            }
        });
    }
    
    // Add keyword warnings
    if (data.suspicious_keywords && data.suspicious_keywords.length > 0) {
        const topKeywords = data.suspicious_keywords.slice(0, 3);
        if (topKeywords.length > 0) {
            flags.push(`🔑 Suspicious keywords: ${topKeywords.join(', ')}`);
        }
    }
    
    // Add link warnings
    if (data.link_analysis && data.link_analysis.length > 0) {
        data.link_analysis.slice(0, 2).forEach(issue => {
            flags.push(`🔗 ${issue}`);
        });
    }

    // Display red flags
    if (flags.length > 0) {
        redFlags.innerHTML = flags.map(flag => `<li>${escapeHtml(flag)}</li>`).join('');
    } else {
        redFlags.innerHTML = '<li style="background: #d1fae5; border-left-color: #10b981;">✅ No major red flags detected</li>';
    }

    // Email details
    emailDetails.innerHTML = `
        <p><strong>Subject:</strong><br>${escapeHtml(data.email_details.subject)}</p>
        <p><strong>From:</strong><br>${escapeHtml(data.email_details.sender)}</p>
        ${data.links && data.links.length > 0 ? 
            `<p><strong>Links Found:</strong> ${data.links.length}</p>` : 
            '<p><strong>Links Found:</strong> None</p>'
        }
    `;

    // Show first link if available
    if (data.links && data.links.length > 0) {
        const firstLink = data.links[0];
        emailDetails.innerHTML += `<div class="link-preview">${escapeHtml(firstLink.url)}</div>`;
    }
}

// Show error message
function showError(message) {
    const error = document.getElementById('error');
    error.textContent = message;
    error.classList.add('active');
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
