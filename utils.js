/* 
 * Advanced Utilities for LLM Agent Demo
 * Additional helper functions and enhancements
 */

// Enhanced Tool Management System
class ToolManager {
    constructor() {
        this.toolMetrics = new Map();
        this.toolHistory = [];
    }

    recordToolExecution(toolName, duration, success, result) {
        if (!this.toolMetrics.has(toolName)) {
            this.toolMetrics.set(toolName, {
                executions: 0,
                totalDuration: 0,
                successCount: 0,
                failureCount: 0
            });
        }

        const metrics = this.toolMetrics.get(toolName);
        metrics.executions++;
        metrics.totalDuration += duration;
        
        if (success) {
            metrics.successCount++;
        } else {
            metrics.failureCount++;
        }

        this.toolHistory.push({
            toolName,
            timestamp: new Date(),
            duration,
            success,
            result: result?.summary || result?.error || 'No result'
        });
    }

    getToolStats(toolName) {
        const metrics = this.toolMetrics.get(toolName);
        if (!metrics) return null;

        return {
            ...metrics,
            averageDuration: metrics.totalDuration / metrics.executions,
            successRate: metrics.successCount / metrics.executions
        };
    }

    getAllStats() {
        const stats = {};
        for (const [toolName, metrics] of this.toolMetrics) {
            stats[toolName] = this.getToolStats(toolName);
        }
        return stats;
    }
}

// Enhanced Response Generator
class ResponseGenerator {
    constructor() {
        this.templates = {
            search: [
                "I found some interesting information about {query}:",
                "Here's what I discovered regarding {query}:",
                "My search for {query} yielded these results:"
            ],
            calculation: [
                "I've performed the calculation for you:",
                "Here's the mathematical result:",
                "The computation gives us:"
            ],
            code: [
                "I've generated and executed the code:",
                "Here's the code solution:",
                "I created this implementation for you:"
            ],
            api: [
                "I retrieved the data from the API:",
                "Here's the information from the external service:",
                "The API call returned:"
            ],
            analysis: [
                "I've analyzed the data and found:",
                "Here are the analytical insights:",
                "The data analysis reveals:"
            ]
        };
    }

    generateIntro(toolType, context = {}) {
        const templates = this.templates[toolType] || this.templates.search;
        const template = templates[Math.floor(Math.random() * templates.length)];
        
        return this.interpolateTemplate(template, context);
    }

    interpolateTemplate(template, context) {
        return template.replace(/\{(\w+)\}/g, (match, key) => {
            return context[key] || match;
        });
    }

    generateSummary(toolResults, originalQuery) {
        const successfulTools = Object.entries(toolResults)
            .filter(([_, result]) => !result.error)
            .length;

        if (successfulTools === 0) {
            return "I encountered some challenges, but I'm ready to try a different approach if you'd like.";
        }

        if (successfulTools === 1) {
            return "I hope this information is helpful! Feel free to ask for more details or a different analysis.";
        }

        return `I used ${successfulTools} different tools to give you a comprehensive answer. Let me know if you need any clarification or additional information!`;
    }
}

// Performance Monitor
class PerformanceMonitor {
    constructor() {
        this.metrics = {
            responseTime: [],
            toolExecutionTime: [],
            memoryUsage: [],
            errorRate: 0,
            totalRequests: 0
        };
    }

    startTimer() {
        return performance.now();
    }

    endTimer(startTime) {
        return performance.now() - startTime;
    }

    recordResponse(duration) {
        this.metrics.responseTime.push(duration);
        this.metrics.totalRequests++;
    }

    recordToolExecution(toolName, duration) {
        this.metrics.toolExecutionTime.push({
            tool: toolName,
            duration: duration,
            timestamp: Date.now()
        });
    }

    recordError() {
        this.metrics.errorRate = (this.metrics.errorRate * (this.metrics.totalRequests - 1) + 1) / this.metrics.totalRequests;
    }

    getAverageResponseTime() {
        if (this.metrics.responseTime.length === 0) return 0;
        return this.metrics.responseTime.reduce((a, b) => a + b, 0) / this.metrics.responseTime.length;
    }

    getPerformanceReport() {
        return {
            averageResponseTime: this.getAverageResponseTime(),
            totalRequests: this.metrics.totalRequests,
            errorRate: this.metrics.errorRate,
            toolExecutions: this.metrics.toolExecutionTime.length,
            recentToolExecutions: this.metrics.toolExecutionTime.slice(-10)
        };
    }
}

// Context Manager for maintaining conversation state
class ContextManager {
    constructor() {
        this.conversationContext = {
            topics: [],
            entities: new Map(),
            previousResults: [],
            userPreferences: new Map()
        };
    }

    addTopicToContext(topic) {
        this.conversationContext.topics.push({
            topic,
            timestamp: Date.now(),
            relevance: 1.0
        });
        
        // Keep only recent topics
        if (this.conversationContext.topics.length > 10) {
            this.conversationContext.topics.shift();
        }
    }

    addEntityToContext(entity, type, metadata = {}) {
        this.conversationContext.entities.set(entity, {
            type,
            metadata,
            mentions: (this.conversationContext.entities.get(entity)?.mentions || 0) + 1,
            lastMentioned: Date.now()
        });
    }

    addResultToContext(result, toolName) {
        this.conversationContext.previousResults.push({
            result,
            toolName,
            timestamp: Date.now()
        });

        // Keep only recent results
        if (this.conversationContext.previousResults.length > 5) {
            this.conversationContext.previousResults.shift();
        }
    }

    getRelevantContext(query) {
        const relevantTopics = this.conversationContext.topics
            .filter(t => this.isTopicRelevant(t.topic, query))
            .slice(-3);

        const relevantEntities = Array.from(this.conversationContext.entities.entries())
            .filter(([entity, _]) => query.toLowerCase().includes(entity.toLowerCase()))
            .slice(0, 5);

        const relevantResults = this.conversationContext.previousResults
            .slice(-2);

        return {
            topics: relevantTopics,
            entities: relevantEntities,
            previousResults: relevantResults
        };
    }

    isTopicRelevant(topic, query) {
        const topicWords = topic.toLowerCase().split(' ');
        const queryWords = query.toLowerCase().split(' ');
        
        return topicWords.some(word => 
            queryWords.some(qWord => 
                qWord.includes(word) || word.includes(qWord)
            )
        );
    }
}

// Enhanced Error Handler
class ErrorHandler {
    constructor() {
        this.errorLog = [];
        this.recoveryStrategies = new Map();
        
        this.setupRecoveryStrategies();
    }

    setupRecoveryStrategies() {
        this.recoveryStrategies.set('network', [
            'Retry the request after a short delay',
            'Use cached data if available',
            'Switch to alternative data source'
        ]);

        this.recoveryStrategies.set('computation', [
            'Simplify the calculation',
            'Break down into smaller steps',
            'Use alternative algorithm'
        ]);

        this.recoveryStrategies.set('parsing', [
            'Clean and sanitize input',
            'Use fuzzy matching',
            'Request clarification from user'
        ]);
    }

    handleError(error, context = {}) {
        const errorInfo = {
            message: error.message,
            type: this.categorizeError(error),
            timestamp: Date.now(),
            context: context
        };

        this.errorLog.push(errorInfo);
        
        // Keep only recent errors
        if (this.errorLog.length > 50) {
            this.errorLog.shift();
        }

        return this.generateErrorResponse(errorInfo);
    }

    categorizeError(error) {
        const message = error.message.toLowerCase();
        
        if (message.includes('network') || message.includes('fetch') || message.includes('timeout')) {
            return 'network';
        }
        
        if (message.includes('calculation') || message.includes('math') || message.includes('expression')) {
            return 'computation';
        }
        
        if (message.includes('parse') || message.includes('invalid') || message.includes('format')) {
            return 'parsing';
        }
        
        return 'general';
    }

    generateErrorResponse(errorInfo) {
        const strategies = this.recoveryStrategies.get(errorInfo.type) || ['Try a different approach'];
        const strategy = strategies[Math.floor(Math.random() * strategies.length)];
        
        return {
            error: `I encountered an issue: ${errorInfo.message}`,
            suggestion: `I suggest we ${strategy.toLowerCase()}.`,
            canRetry: ['network', 'computation'].includes(errorInfo.type)
        };
    }

    getErrorStats() {
        const recentErrors = this.errorLog.slice(-20);
        const errorTypes = {};
        
        recentErrors.forEach(error => {
            errorTypes[error.type] = (errorTypes[error.type] || 0) + 1;
        });

        return {
            totalErrors: this.errorLog.length,
            recentErrors: recentErrors.length,
            errorTypes: errorTypes,
            lastError: this.errorLog[this.errorLog.length - 1]
        };
    }
}

// Export utilities for use in main agent
if (typeof window !== 'undefined') {
    window.AgentUtils = {
        ToolManager,
        ResponseGenerator,
        PerformanceMonitor,
        ContextManager,
        ErrorHandler
    };
}
