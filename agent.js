class LLMAgent {
    constructor() {
        this.tools = {
            webSearch: new WebSearchTool(),
            calculator: new CalculatorTool(),
            codeExecutor: new CodeExecutorTool(),
            apiCaller: new APICallerTool(),
            dataAnalyzer: new DataAnalyzerTool()
        };
        
        this.conversationHistory = [];
        this.isThinking = false;
        
        this.initializeUI();
    }

    initializeUI() {
        const sendButton = document.getElementById('sendButton');
        const userInput = document.getElementById('userInput');
        
        sendButton.addEventListener('click', () => this.handleUserInput());
        userInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                this.handleUserInput();
            }
        });
    }

    async handleUserInput() {
        const userInput = document.getElementById('userInput');
        const message = userInput.value.trim();
        
        if (!message || this.isThinking) return;
        
        userInput.value = '';
        this.addMessage('user', message);
        
        this.setStatus('thinking', 'Processing...');
        this.isThinking = true;
        
        try {
            await this.processMessage(message);
        } catch (error) {
            console.error('Error processing message:', error);
            this.addMessage('agent', '❌ Sorry, I encountered an error while processing your request. Please try again.');
            this.setStatus('error', 'Error');
        } finally {
            this.isThinking = false;
            this.setStatus('ready', 'Ready');
        }
    }

    async processMessage(message) {
        // Simulate LLM reasoning about which tools to use
        const plan = this.analyzePlan(message);
        
        if (plan.tools.length === 0) {
            // Simple response without tools
            const response = this.generateSimpleResponse(message);
            this.addMessage('agent', response);
            return;
        }

        // Execute plan with tools
        let context = { originalQuery: message, results: {} };
        
        for (const toolPlan of plan.tools) {
            try {
                this.setToolActive(toolPlan.tool);
                const result = await this.executeTool(toolPlan.tool, toolPlan.params, context);
                context.results[toolPlan.tool] = result;
                this.setToolIdle(toolPlan.tool);
                
                // Add a small delay for better UX
                await this.sleep(500);
            } catch (error) {
                console.error(`Tool ${toolPlan.tool} failed:`, error);
                context.results[toolPlan.tool] = { error: error.message };
                this.setToolIdle(toolPlan.tool);
            }
        }

        // Generate final response based on tool results
        const finalResponse = this.synthesizeResponse(message, context.results, plan);
        this.addMessage('agent', finalResponse);
    }

    analyzePlan(message) {
        const lowerMessage = message.toLowerCase();
        const plan = { tools: [], reasoning: '' };

        // Simple rule-based planning (in a real implementation, this would be LLM-powered)
        if (lowerMessage.includes('search') || lowerMessage.includes('find') || lowerMessage.includes('news') || lowerMessage.includes('current') || lowerMessage.includes('latest')) {
            plan.tools.push({ tool: 'webSearch', params: { query: message } });
        }

        if (lowerMessage.includes('calculate') || lowerMessage.includes('math') || /\d+.*[\+\-\*\/].*\d+/.test(message)) {
            plan.tools.push({ tool: 'calculator', params: { expression: message } });
        }

        if (lowerMessage.includes('code') || lowerMessage.includes('program') || lowerMessage.includes('chart') || lowerMessage.includes('graph') || lowerMessage.includes('visualize')) {
            plan.tools.push({ tool: 'codeExecutor', params: { intent: message } });
        }

        if (lowerMessage.includes('api') || lowerMessage.includes('weather') || lowerMessage.includes('stock') || lowerMessage.includes('currency')) {
            plan.tools.push({ tool: 'apiCaller', params: { intent: message } });
        }

        if (lowerMessage.includes('analyze') || lowerMessage.includes('data') || lowerMessage.includes('statistics') || lowerMessage.includes('trend')) {
            plan.tools.push({ tool: 'dataAnalyzer', params: { intent: message } });
        }

        return plan;
    }

    async executeTool(toolName, params, context) {
        const tool = this.tools[toolName];
        if (!tool) {
            throw new Error(`Tool ${toolName} not found`);
        }

        this.logExecution(`🔧 Executing ${toolName} with params:`, params);
        const result = await tool.execute(params, context);
        this.logExecution(`✅ ${toolName} completed:`, result);
        
        return result;
    }

    synthesizeResponse(originalQuery, toolResults, plan) {
        let response = "Based on your request, I've used the following tools:\n\n";

        Object.entries(toolResults).forEach(([toolName, result]) => {
            if (result.error) {
                response += `❌ **${this.getToolDisplayName(toolName)}**: ${result.error}\n\n`;
            } else {
                response += `✅ **${this.getToolDisplayName(toolName)}**: ${result.summary || 'Completed successfully'}\n\n`;
                
                if (result.data) {
                    response += `${result.data}\n\n`;
                }
            }
        });

        // Add conclusion
        response += "---\n\n";
        response += this.generateConclusion(originalQuery, toolResults);

        return response;
    }

    generateConclusion(query, results) {
        // Simple conclusion generation (in a real implementation, this would be LLM-powered)
        const successfulTools = Object.entries(results).filter(([_, result]) => !result.error);
        
        if (successfulTools.length === 0) {
            return "I encountered some issues with the tools, but I'm ready to try again with a different approach.";
        }

        return "I hope this information helps! Feel free to ask follow-up questions or request additional analysis.";
    }

    generateSimpleResponse(message) {
        // Simple responses for queries that don't need tools
        const responses = [
            "I understand your question. Could you be more specific about what kind of information or action you need?",
            "That's an interesting question! To provide the best help, could you clarify what you'd like me to search for, calculate, or analyze?",
            "I'm here to help! I can search the web, perform calculations, execute code, call APIs, or analyze data. What would you like me to do?"
        ];
        
        return responses[Math.floor(Math.random() * responses.length)];
    }

    getToolDisplayName(toolName) {
        const names = {
            webSearch: 'Web Search',
            calculator: 'Calculator',
            codeExecutor: 'Code Executor',
            apiCaller: 'API Caller',
            dataAnalyzer: 'Data Analyzer'
        };
        return names[toolName] || toolName;
    }

    addMessage(type, content) {
        const chatMessages = document.getElementById('chatMessages');
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}-message`;
        
        const contentDiv = document.createElement('div');
        contentDiv.className = 'message-content';
        
        if (type === 'agent') {
            contentDiv.innerHTML = marked.parse(content);
        } else {
            contentDiv.innerHTML = `<strong>${type === 'user' ? 'You' : 'Agent'}:</strong> ${content}`;
        }
        
        messageDiv.appendChild(contentDiv);
        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    setStatus(type, message) {
        const statusElement = document.getElementById('status');
        const statusDot = document.getElementById('statusDot');
        
        statusElement.textContent = message;
        statusDot.className = `status-dot ${type}`;
    }

    setToolActive(toolName) {
        const toolCard = document.getElementById(`${toolName}Tool`);
        if (toolCard) {
            toolCard.classList.add('active');
            const statusElement = toolCard.querySelector('.tool-status');
            if (statusElement) {
                statusElement.textContent = 'Active';
            }
        }
    }

    setToolIdle(toolName) {
        const toolCard = document.getElementById(`${toolName}Tool`);
        if (toolCard) {
            toolCard.classList.remove('active');
            const statusElement = toolCard.querySelector('.tool-status');
            if (statusElement) {
                statusElement.textContent = 'Idle';
            }
        }
    }

    logExecution(message, data) {
        const executionResults = document.getElementById('executionResults');
        const timestamp = new Date().toLocaleTimeString();
        executionResults.textContent += `[${timestamp}] ${message}\n${JSON.stringify(data, null, 2)}\n\n`;
        executionResults.scrollTop = executionResults.scrollHeight;
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Tool implementations
class WebSearchTool {
    async execute(params, context) {
        // Simulate web search (in a real implementation, this would call a search API)
        await this.delay(1000 + Math.random() * 2000);
        
        const mockResults = [
            "Recent developments in AI show significant progress in multimodal capabilities.",
            "New research indicates improved efficiency in large language models.",
            "Industry leaders are focusing on responsible AI development and deployment.",
            "Breakthrough in neural architecture design shows promise for edge computing.",
            "Latest studies reveal enhanced reasoning capabilities in transformer models."
        ];
        
        const randomResult = mockResults[Math.floor(Math.random() * mockResults.length)];
        
        return {
            summary: "Found relevant information from web search",
            data: `**Search Results for "${params.query}":**\n\n${randomResult}\n\n*Note: This is a demo with simulated results.*`,
            sources: ["example.com", "research.ai", "tech-news.com"]
        };
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

class CalculatorTool {
    async execute(params, context) {
        await this.delay(500);
        
        try {
            // Extract mathematical expressions from the input
            const expression = this.extractMathExpression(params.expression);
            
            if (!expression) {
                return {
                    summary: "No mathematical expression found",
                    data: "I couldn't find a mathematical expression to calculate. Please provide a clear mathematical operation."
                };
            }

            // Safe evaluation (limited to basic math)
            const result = this.safeEval(expression);
            
            return {
                summary: `Calculated: ${expression} = ${result}`,
                data: `**Calculation Result:**\n\n${expression} = **${result}**`
            };
        } catch (error) {
            return {
                error: `Calculation failed: ${error.message}`
            };
        }
    }

    extractMathExpression(text) {
        // Simple regex to extract basic math expressions
        const mathPattern = /(\d+(?:\.\d+)?)\s*([+\-*/])\s*(\d+(?:\.\d+)?)/g;
        const match = mathPattern.exec(text);
        
        if (match) {
            return `${match[1]} ${match[2]} ${match[3]}`;
        }
        
        // Look for simple numbers for operations like square, factorial, etc.
        const numberPattern = /(\d+(?:\.\d+)?)/;
        const numberMatch = numberPattern.exec(text);
        
        if (numberMatch && text.toLowerCase().includes('square')) {
            return `${numberMatch[1]} ^ 2`;
        }
        
        return null;
    }

    safeEval(expression) {
        // Very basic and safe math evaluation
        const sanitized = expression.replace(/[^0-9+\-*/.() ]/g, '');
        
        if (expression.includes('^')) {
            const parts = expression.split('^');
            return Math.pow(parseFloat(parts[0].trim()), parseFloat(parts[1].trim()));
        }
        
        try {
            return Function(`"use strict"; return (${sanitized})`)();
        } catch (error) {
            throw new Error('Invalid mathematical expression');
        }
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

class CodeExecutorTool {
    async execute(params, context) {
        await this.delay(1000);
        
        try {
            const code = this.generateCode(params.intent);
            
            // Create a safe execution environment
            const result = this.executeCode(code);
            
            return {
                summary: "Code executed successfully",
                data: `**Generated and Executed Code:**\n\n\`\`\`javascript\n${code}\n\`\`\`\n\n**Output:**\n${result}`
            };
        } catch (error) {
            return {
                error: `Code execution failed: ${error.message}`
            };
        }
    }

    generateCode(intent) {
        const lowerIntent = intent.toLowerCase();
        
        if (lowerIntent.includes('prime') || lowerIntent.includes('prime numbers')) {
            return `
// Generate prime numbers
function generatePrimes(limit) {
    const primes = [];
    for (let num = 2; num <= limit; num++) {
        let isPrime = true;
        for (let i = 2; i <= Math.sqrt(num); i++) {
            if (num % i === 0) {
                isPrime = false;
                break;
            }
        }
        if (isPrime) primes.push(num);
    }
    return primes;
}

const primes = generatePrimes(50);
console.log("First 15 prime numbers:", primes.slice(0, 15));
`;
        } else if (lowerIntent.includes('fibonacci')) {
            return `
// Generate Fibonacci sequence
function fibonacci(n) {
    const fib = [0, 1];
    for (let i = 2; i < n; i++) {
        fib[i] = fib[i-1] + fib[i-2];
    }
    return fib;
}

const fibSequence = fibonacci(15);
console.log("Fibonacci sequence (15 numbers):", fibSequence);
`;
        } else if (lowerIntent.includes('chart') || lowerIntent.includes('graph')) {
            return `
// Generate sample data visualization
const data = Array.from({length: 10}, (_, i) => ({
    x: i + 1,
    y: Math.floor(Math.random() * 100) + 1
}));

console.log("Sample chart data:", data);
console.log("Chart would display:", data.map(d => \`(\${d.x}, \${d.y})\`).join(', '));
`;
        } else {
            return `
// General purpose code example
const currentTime = new Date().toLocaleString();
const randomNumber = Math.floor(Math.random() * 1000);

console.log("Current time:", currentTime);
console.log("Random number:", randomNumber);
console.log("Your request:", "${intent}");
`;
        }
    }

    executeCode(code) {
        // Capture console.log output
        const logs = [];
        const originalLog = console.log;
        console.log = (...args) => {
            logs.push(args.join(' '));
        };

        try {
            // Execute the code in a safe way
            eval(code);
            console.log = originalLog;
            return logs.join('\n');
        } catch (error) {
            console.log = originalLog;
            throw error;
        }
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

class APICallerTool {
    async execute(params, context) {
        await this.delay(1500);
        
        const intent = params.intent.toLowerCase();
        
        if (intent.includes('weather')) {
            return this.mockWeatherAPI();
        } else if (intent.includes('stock') || intent.includes('price')) {
            return this.mockStockAPI();
        } else if (intent.includes('currency') || intent.includes('exchange')) {
            return this.mockCurrencyAPI();
        } else {
            return this.mockGenericAPI();
        }
    }

    mockWeatherAPI() {
        const cities = ['New York', 'London', 'Tokyo', 'Sydney', 'Paris'];
        const city = cities[Math.floor(Math.random() * cities.length)];
        const temp = Math.floor(Math.random() * 30) + 5;
        const conditions = ['Sunny', 'Cloudy', 'Rainy', 'Partly Cloudy'];
        const condition = conditions[Math.floor(Math.random() * conditions.length)];
        
        return {
            summary: `Weather data retrieved for ${city}`,
            data: `**Weather Information:**\n\n🌍 **Location:** ${city}\n🌡️ **Temperature:** ${temp}°C\n☁️ **Conditions:** ${condition}\n\n*Note: This is demo data.*`
        };
    }

    mockStockAPI() {
        const stocks = ['AAPL', 'GOOGL', 'MSFT', 'TSLA', 'AMZN'];
        const stock = stocks[Math.floor(Math.random() * stocks.length)];
        const price = (Math.random() * 1000 + 50).toFixed(2);
        const change = (Math.random() * 10 - 5).toFixed(2);
        
        return {
            summary: `Stock price retrieved for ${stock}`,
            data: `**Stock Information:**\n\n📈 **Symbol:** ${stock}\n💰 **Price:** $${price}\n📊 **Change:** ${change > 0 ? '+' : ''}${change}\n\n*Note: This is demo data.*`
        };
    }

    mockCurrencyAPI() {
        const currencies = [
            { from: 'USD', to: 'EUR', rate: 0.85 },
            { from: 'USD', to: 'GBP', rate: 0.73 },
            { from: 'USD', to: 'JPY', rate: 110.25 },
            { from: 'EUR', to: 'USD', rate: 1.18 }
        ];
        
        const currency = currencies[Math.floor(Math.random() * currencies.length)];
        
        return {
            summary: `Exchange rate retrieved`,
            data: `**Currency Exchange:**\n\n💱 **Pair:** ${currency.from}/${currency.to}\n📈 **Rate:** ${currency.rate}\n\n*Note: This is demo data.*`
        };
    }

    mockGenericAPI() {
        return {
            summary: "API call completed",
            data: "**API Response:**\n\nGeneric API endpoint called successfully. In a real implementation, this would connect to actual external APIs.\n\n*Note: This is a demonstration.*"
        };
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

class DataAnalyzerTool {
    async execute(params, context) {
        await this.delay(1200);
        
        try {
            const analysis = this.performAnalysis(params.intent, context);
            
            return {
                summary: "Data analysis completed",
                data: analysis
            };
        } catch (error) {
            return {
                error: `Data analysis failed: ${error.message}`
            };
        }
    }

    performAnalysis(intent, context) {
        // Generate sample data for analysis
        const data = this.generateSampleData();
        
        const stats = {
            count: data.length,
            mean: this.calculateMean(data),
            median: this.calculateMedian(data),
            stdDev: this.calculateStdDev(data),
            min: Math.min(...data),
            max: Math.max(...data)
        };

        return `**Data Analysis Results:**

📊 **Dataset:** Sample data (${stats.count} points)
📈 **Mean:** ${stats.mean.toFixed(2)}
📊 **Median:** ${stats.median}
📉 **Standard Deviation:** ${stats.stdDev.toFixed(2)}
🔻 **Minimum:** ${stats.min}
🔺 **Maximum:** ${stats.max}

**Sample Data:** ${data.slice(0, 10).join(', ')}${data.length > 10 ? '...' : ''}

**Insights:**
- The data shows ${stats.mean > stats.median ? 'a right-skewed' : 'a left-skewed'} distribution
- Range: ${stats.max - stats.min}
- The data appears to be ${stats.stdDev < stats.mean / 3 ? 'relatively stable' : 'quite variable'}

*Note: This is demo analysis with generated data.*`;
    }

    generateSampleData() {
        const size = 50 + Math.floor(Math.random() * 50);
        return Array.from({length: size}, () => Math.floor(Math.random() * 100) + 1);
    }

    calculateMean(data) {
        return data.reduce((sum, val) => sum + val, 0) / data.length;
    }

    calculateMedian(data) {
        const sorted = [...data].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        return sorted.length % 2 === 0 
            ? (sorted[mid - 1] + sorted[mid]) / 2 
            : sorted[mid];
    }

    calculateStdDev(data) {
        const mean = this.calculateMean(data);
        const variance = data.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / data.length;
        return Math.sqrt(variance);
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Initialize the agent when the page loads
document.addEventListener('DOMContentLoaded', () => {
    window.agent = new LLMAgent();
});
