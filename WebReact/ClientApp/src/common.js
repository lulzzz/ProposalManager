Object.defineProperty(exports, "__esModule", { value: true });
exports.oppStatus = ['Not Started',
    'In Progress',
    'Blocked',
    'Completed'
];

/* Dashboard list */
exports.oppStatusText = [
    'None Empty',
    'Creating',
    'In Progress',
    'Assigned',
    'Draft',
    'Not Started',
    'In Review',
    'Blocked',
    'Completed',
    'Submitted',
    'Accepted'
];

exports.oppStatusClassName = [
    'NoneEmpty',
    'Creating',
    'InProgress',
    'Assigned',
    'Draft',
    'NotStarted',
    'InReview',
    'Blocked',
    'Completed',
    'Submitted',
    'Accepted'
];

exports.oppStatusTextOld = [{
    'NotStarted': 'Not Started',
    'InProgress': 'In Progress',
    'Blocked': 'Blocked',
    'Completed': 'Completed'
}];

exports.channels = [
    {
        name: "Risk Assessment",
        description: "Risk Assessment channel"
    },
    {
        name: "Credit Check",
        description: "Credit Check channel"
    },
    {
        name: "Compliance",
        description: "Compliance channel"
    },
    {
        name: "Formal Proposal",
        description: "Formal Proposal channel"
    },
    {
        name: "Customer Decision",
        description: "Customer Decision channel"
    }
];



exports.userRoles = ['Loan Officer', 'Relationship Manager', 'Credit Analyst', 'Legal Counsel', 'Senior Risk Officer'];

// Get the value of query parameter
exports.getQueryVariable = (variable) => {
    const query = window.location.search.substring(1);
    const vars = query.split('&');
    for (const varPairs of vars) {
        const pair = varPairs.split('=');
        if (decodeURIComponent(pair[0]) === variable) {
            return decodeURIComponent(pair[1]);
        }
    }
    return null;
};