/**
 * Improved search emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const {
  clampCount,
  buildMessageFilter,
  matchesMessageFilter,
  formatEmailList,
  UNTRUSTED_NOTE
} = require('./format');

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args = {}) {
  const folder = args.folder || "inbox";
  const requestedCount = clampCount(args.count, 10);
  const strict = args.strict === true;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  const filterTerms = {
    hasAttachments,
    unreadOnly,
    receivedAfter: args.receivedAfter,
    receivedBefore: args.receivedBefore,
    flagStatus: args.flagStatus
  };

  try {
    // Validate filters (dates, flagStatus) before touching the API
    buildMessageFilter(filterTerms);

    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);
    
    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint, 
      accessToken, 
      { query, from, to, subject },
      filterTerms,
      requestedCount,
      { strict }
    );
    
    return formatSearchResults(response);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    // General error response
    return {
      content: [{ 
        type: "text", 
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount, options = {}) {
  // Track search strategies attempted
  const searchAttempts = [];
  const finish = (response, strategy) => {
    response.value = (response.value || []).filter(email => matchesMessageFilter(email, filterTerms));
    response._searchInfo = {
      attemptsCount: searchAttempts.length,
      strategies: searchAttempts,
      strategy,
      // true when the results came from a broader search than asked for
      fallback: strategy !== 'combined-search',
      originalTerms: searchTerms,
      filterTerms
    };
    return response;
  };
  
  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount));
    if (config.DEBUG) console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");
    
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
    if ((response.value && response.value.length > 0) || options.strict) {
      console.error(`Combined search found ${response.value?.length || 0} results`);
      return finish(response, 'combined-search');
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
    if (options.strict) throw error;
  }
  
  // 2. Try each search term individually, starting with most specific
  const searchPriority = ['subject', 'from', 'to', 'query'];
  
  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}`);
        searchAttempts.push(`single-term-${term}`);
        
        // For single term search, only use $search with that term
        // Graph API does not support $orderby or $filter with $search
        const simplifiedParams = {
          $top: Math.min(50, maxCount),
          $select: config.EMAIL_SELECT_FIELDS
        };
        
        // Build KQL terms for search
        const kqlParts = [];
        
        // Add the search term in the appropriate KQL syntax
        if (term === 'query') {
          // General query doesn't need a prefix
          kqlParts.push(searchTerms[term]);
        } else {
          // Specific field searches use field:value syntax
          kqlParts.push(kqlField(term, searchTerms[term]));
        }
        
        // Add boolean filters as KQL (can't use $filter with $search)
        addBooleanFiltersAsKQL(kqlParts, filterTerms);
        
        simplifiedParams.$search = `"${kqlParts.join(' ')}"`;
        
        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
        if (response.value && response.value.length > 0) {
          console.error(`Search with ${term} successful: found ${response.value.length} results`);
          return finish(response, `single-term-${term}`);
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }
  
  // 3. Try with only the filters (dates, unread, attachments, flag)
  if (buildMessageFilter(filterTerms)) {
    try {
      console.error("Attempting search with only boolean filters");
      searchAttempts.push("boolean-filters-only");
      
      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: config.EMAIL_SELECT_FIELDS,
        $orderby: 'receivedDateTime desc'
      };
      
      filterOnlyParams.$filter = buildMessageFilter(filterTerms);

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Filter-only search found ${response.value?.length || 0} results`);
      return finish(response, 'boolean-filters-only');
    } catch (error) {
      console.error(`Boolean filter search failed: ${error.message}`);
    }
  }
  
  // 4. Final fallback: just get recent emails with pagination
  console.error("All search strategies failed, falling back to recent emails");
  searchAttempts.push("recent-emails");
  
  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  
  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
  console.error(`Fallback to recent emails found ${response.value?.length || 0} results`);

  // Deliberately not filtered: this is "recent mail", flagged as such below
  response._searchInfo = {
    attemptsCount: searchAttempts.length,
    strategies: searchAttempts,
    strategy: 'recent-emails',
    fallback: true,
    originalTerms: searchTerms,
    filterTerms
  };
  return response;
}

/**
 * A KQL field restriction for use inside the double-quoted $search value.
 * Inner quotes must be backslash-escaped: `"from:"x""` is a Graph syntax
 * error, `"from:\\"x\\""` works and allows values with spaces.
 * @param {string} field - KQL property (subject, from, to)
 * @param {string} value - Value to match
 * @returns {string}
 */
function kqlField(field, value) {
  const clean = String(value).replace(/["\\]/g, '');
  return `${field}:\\"${clean}\\"`;
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} count - Maximum number of results
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };
  
  // Handle search terms
  const kqlTerms = [];
  
  if (searchTerms.query) {
    // General query doesn't need a prefix
    kqlTerms.push(searchTerms.query);
  }
  
  if (searchTerms.subject) {
    kqlTerms.push(kqlField('subject', searchTerms.subject));
  }

  if (searchTerms.from) {
    kqlTerms.push(kqlField('from', searchTerms.from));
  }

  if (searchTerms.to) {
    kqlTerms.push(kqlField('to', searchTerms.to));
  }
  
  // Add $search if we have any search terms
  if (kqlTerms.length > 0) {
    // Graph API does not support $orderby or $filter with $search
    // Move boolean filters into KQL syntax instead
    addBooleanFiltersAsKQL(kqlTerms, filterTerms);
    params.$search = `"${kqlTerms.join(' ')}"`;
  } else {
    // No search terms — safe to use $orderby and $filter
    params.$orderby = 'receivedDateTime desc';
    const filter = buildMessageFilter(filterTerms);
    if (filter) params.$filter = filter;
  }
  
  return params;
}

/**
 * Add boolean filters as KQL terms for use with $search
 * Use this instead of addBooleanFilters when $search is present
 * @param {string[]} kqlTerms - Array of KQL terms to append to
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFiltersAsKQL(kqlTerms, filterTerms) {
  if (filterTerms.hasAttachments === true) {
    kqlTerms.push('hasAttachments:true');
  }
  
  if (filterTerms.unreadOnly === true) {
    kqlTerms.push('isRead:false');
  }

  // Day granularity in KQL; exact times and flag status are re-checked client-side
  if (filterTerms.receivedAfter) {
    kqlTerms.push(`received>=${new Date(filterTerms.receivedAfter).toISOString().slice(0, 10)}`);
  }
  if (filterTerms.receivedBefore) {
    kqlTerms.push(`received<=${new Date(filterTerms.receivedBefore).toISOString().slice(0, 10)}`);
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @returns {object} - MCP response object
 */
function formatSearchResults(response) {
  const info = response._searchInfo;

  if (!response.value || response.value.length === 0) {
    return {
      content: [{
        type: "text",
        text: `No emails found matching your search criteria.`
      }]
    };
  }

  const emailList = formatEmailList(response.value);

  let header = `Found ${response.value.length} emails matching your search criteria:`;
  if (info && info.fallback) {
    header = `WARNING: no emails matched all of your search criteria. ` +
      `Showing ${response.value.length} results from a broader "${info.strategy}" search instead; ` +
      `these may NOT match what you asked for. Pass strict: true to get only exact matches.`;
  } else if (info && info.strategy) {
    header += `\n(Search used ${info.strategy} strategy)`;
  }

  return {
    content: [{
      type: "text",
      text: `${header}\n${UNTRUSTED_NOTE}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
module.exports.buildSearchParams = buildSearchParams;
module.exports.kqlField = kqlField;
