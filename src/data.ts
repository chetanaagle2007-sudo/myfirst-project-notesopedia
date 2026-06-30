import { StudyUnit, ExpectedQuestion, VivaQuestion, MemoryTrick } from "./types";

export const STUDY_UNITS: StudyUnit[] = [
  {
    id: "unit-1-intro",
    title: "Unit 1: Introduction & Database Design",
    description: "Features of popular RDBMS, DBMS vs RDBMS, Normalization (1NF, 2NF, 3NF), and constraints.",
    topics: [
      {
        id: "rdbms-intro",
        title: "Relational Database Management Systems (RDBMS)",
        definition: "An RDBMS is a database system based on the relational model introduced by E.F. Codd. It stores data in tabular structures consisting of rows (tuples) and columns (attributes) and maintains integrity constraints across relationships.",
        typesOrFeatures: [
          "Tables (Relations): The fundamental logical structure where data is stored in row-column format.",
          "Primary and Foreign Keys: Keys used to uniquely identify records and establish relations between tables.",
          "Data Integrity Constraints: Rules enforced at the database level (Entity, Domain, and Referential Integrity).",
          "ACID Transactions: Guarantees reliable transaction execution even during system crashes."
        ],
        advantages: [
          "Elimination of Data Redundancy: Through normalization techniques.",
          "Referential Integrity: Prevents orphaned records and maintains data consistency across multiple entities.",
          "High Security: Rich access control models with schemas, roles, and privileges.",
          "Standardized Querying: Handled globally through SQL (Structured Query Language)."
        ],
        example: `-- Creating a standard relational table structure with keys
CREATE TABLE departments (
    dept_id SERIAL PRIMARY KEY,
    dept_name VARCHAR(100) NOT NULL UNIQUE
);

CREATE TABLE employees (
    emp_id SERIAL PRIMARY KEY,
    emp_name VARCHAR(100) NOT NULL,
    dept_id INT REFERENCES departments(dept_id) ON DELETE CASCADE
);`
      },
      {
        id: "dbms-vs-rdbms",
        title: "Difference Between DBMS and RDBMS",
        definition: "A DBMS stores data as files and does not enforce relational associations between files, whereas an RDBMS stores data in tabular schemas and strictly enforces structural constraints (like primary keys and foreign keys) with relational links.",
        typesOrFeatures: [
          "Data Storage: DBMS uses file-based or XML hierarchies; RDBMS uses structured tables.",
          "Integrity Constraints: DBMS doesn't enforce referential integrity at the engine level; RDBMS strictly enforces it.",
          "Normalization: DBMS rarely supports or guides normalization; RDBMS is built around normal forms (1NF, 2NF, 3NF).",
          "Distributed Databases: DBMS generally runs on single machines; RDBMS is highly distributed and scalable."
        ],
        advantages: [
          "Better Security: RDBMS offers granular access control down to columns and rows.",
          "No Redundancy: Relational model enforces normalized structures.",
          "High Concurrency: Support for advanced multi-user locking protocols."
        ],
        example: `-- Under RDBMS, this referential link ensures no invalid dept_id is ever inserted.
-- A simple DBMS would allow putting any garbage integer in the dept_id field.
INSERT INTO employees (emp_name, dept_id) VALUES ('Alice Smith', 9999);
-- ERROR: insert or update on table "employees" violates foreign key constraint`
      },
      {
        id: "normalization",
        title: "Normalization (1NF, 2NF, 3NF)",
        definition: "Normalization is a systematic database design technique used to organize tables to minimize data redundancy, avoid insertion/deletion/update anomalies, and enforce relational dependency integrity.",
        typesOrFeatures: [
          "1NF (First Normal Form): Every column must contain atomic (indivisible) values, and there must be no repeating groups.",
          "2NF (Second Normal Form): Must be in 1NF, and all non-key attributes must be fully functionally dependent on the entire primary key (no partial dependency).",
          "3NF (Third Normal Form): Must be in 2NF, and there must be no transitive dependencies (non-key attributes must not depend on other non-key attributes)."
        ],
        advantages: [
          "Reduces redundant storage space.",
          "Eliminates data anomalies (Insertion, Deletion, Update).",
          "Simplifies database maintenance and schema updates."
        ],
        example: `-- 1NF Violating: (Hobbies: 'Coding, Reading')
-- 1NF Corrected: Separate row or table for hobbies.

-- 2NF Violation: composite key (StudentID, CourseID) -> CourseFee depends only on CourseID.
-- 2NF Corrected: Split into StudentCourse(StudentID, CourseID) and Course(CourseID, CourseFee).

-- 3NF Violation: Employee(EmpID -> DeptID -> DeptLocation) - Transitive dependency!
-- 3NF Corrected: Split into Employee(EmpID, EmpName, DeptID) and Department(DeptID, DeptLocation).`
      }
    ]
  },
  {
    id: "unit-2-plsql",
    title: "Unit 2: PL/SQL & Procedural Features",
    description: "PL/SQL block structures, data types (%type, %rowtype), cursor types, triggers, procedures, functions, packages, and exception handling.",
    topics: [
      {
        id: "plsql-overview",
        title: "PL/SQL & Procedural Blocks",
        definition: "PL/pgSQL (or PL/SQL in general) is a procedural language extension of SQL. It combines the data-manipulation power of SQL with the control structures (loops, conditionals, variables) of procedural programming.",
        syntax: `DECLARE
    -- Variable declarations (formal parameters, %TYPE, %ROWTYPE)
BEGIN
    -- Sequential statements, Loops, Conditionals, SQL
EXCEPTION
    -- Exception handlers
END;`,
        typesOrFeatures: [
          "Anonymous Blocks: Unnamed blocks compiled and executed dynamically at runtime.",
          "Named Blocks: Stored routines (Procedures, Functions, Triggers, Packages) saved in the database catalog."
        ],
        advantages: [
          "Reduced Network Traffic: Sends a bundle of statements at once instead of individual queries.",
          "Better Performance: Compiled server-side for speed.",
          "Procedural Power: Supports structured logic, loops, conditional branching, and error safety inside DB."
        ],
        example: `DO $$
DECLARE
    emp_count INTEGER;
BEGIN
    SELECT COUNT(*) INTO emp_count FROM employees;
    RAISE NOTICE 'Total number of employees: %', emp_count;
END $$;`
      },
      {
        id: "type-rowtype-record",
        title: "%TYPE, %ROWTYPE, and RECORD",
        definition: "These are dynamic attribute declarations in PL/SQL used to bind a variable's data type directly to an existing database column or table structure. This ensures type safety and compatibility when the database schema changes.",
        typesOrFeatures: [
          "%TYPE: Sets variable type to the exact data type of a specific table column.",
          "%ROWTYPE: Sets a variable as a row-record container representing an entire row of a table.",
          "RECORD: A custom composite data type whose structure can change dynamically based on the query result assigned to it."
        ],
        advantages: [
          "Schema Independence: If you change a column from VARCHAR(50) to VARCHAR(100), the PL/SQL code auto-updates.",
          "Type Safety: Guarantees no truncation or type mismatch errors during variable assignment.",
          "Highly readable and compact variable lists."
        ],
        example: `CREATE OR REPLACE FUNCTION get_employee_details(p_emp_id INT)
RETURNS VOID AS $$
DECLARE
    -- %TYPE binds to employees.emp_name
    v_name employees.emp_name%TYPE; 
    -- %ROWTYPE binds to the whole employees structure
    v_row employees%ROWTYPE; 
BEGIN
    SELECT emp_name INTO v_name FROM employees WHERE emp_id = p_emp_id;
    SELECT * INTO v_row FROM employees WHERE emp_id = p_emp_id;
    
    RAISE NOTICE 'Name is %, and row name is %', v_name, v_row.emp_name;
END;
$$ LANGUAGE plpgsql;`
      },
      {
        id: "cursors",
        title: "Cursors (Implicit, Explicit, Parameterized)",
        definition: "A Cursor is a pointer to the private SQL context area allocated by the RDBMS to process and iterate through individual rows of a multi-row SELECT query result set.",
        syntax: `DECLARE
    cursor_name CURSOR FOR select_statement;
BEGIN
    OPEN cursor_name;
    FETCH cursor_name INTO variables;
    CLOSE cursor_name;
END;`,
        typesOrFeatures: [
          "Implicit Cursors: Automatically opened and handled by the RDBMS engine for single-row queries (SELECT INTO, INSERT, UPDATE, DELETE). Attributes: %FOUND, %NOTFOUND, %ROWCOUNT, %ISOPEN.",
          "Explicit Cursors: Declared, opened, fetched, and closed manually by the developer for processing multi-row results.",
          "Parameterized Cursors: Explicit cursors that accept input parameters at the time of opening to filter row results dynamically."
        ],
        advantages: [
          "Efficient memory usage by processing one row at a time instead of fetching millions into application RAM.",
          "Enables granular row-by-row procedural operations and updates."
        ],
        example: `CREATE OR REPLACE FUNCTION process_dept_salary(p_dept_id INT) 
RETURNS VOID AS $$
DECLARE
    v_emp_id INT;
    v_salary NUMERIC;
    -- Parameterized Explicit Cursor
    emp_cur CURSOR(d_id INT) FOR 
        SELECT emp_id, salary FROM employees WHERE dept_id = d_id;
BEGIN
    OPEN emp_cur(p_dept_id);
    LOOP
        FETCH emp_cur INTO v_emp_id, v_salary;
        EXIT WHEN NOT FOUND; -- Explicit cursor attribute check
        
        -- Perform row-by-row logic
        UPDATE employees SET salary = v_salary * 1.1 WHERE emp_id = v_emp_id;
    END LOOP;
    CLOSE emp_cur;
END;
$$ LANGUAGE plpgsql;`
      },
      {
        id: "triggers",
        title: "Triggers",
        definition: "A Trigger is a specialized stored program that automatically executes (fires) in response to a specific event (INSERT, UPDATE, DELETE) on a particular table or view.",
        syntax: `CREATE TRIGGER trigger_name
{BEFORE | AFTER | INSTEAD OF} {INSERT | UPDATE | DELETE}
ON table_name
[FOR EACH ROW]
EXECUTE FUNCTION function_name();`,
        typesOrFeatures: [
          "Row-Level Triggers (FOR EACH ROW): Fires once for every single affected row. Excellent for audits/calculations.",
          "Statement-Level Triggers: Fires once per SQL statement, regardless of how many rows are affected.",
          "BEFORE / AFTER: Determines whether the trigger runs prior to or following the actual SQL operation.",
          "INSTEAD OF: Typically written on complex views to redirect modifications to base tables."
        ],
        advantages: [
          "Enforces complex business rules and auditing that constraints cannot handle.",
          "Automatically logs database activity history.",
          "Maintains derived summary data synchronously."
        ],
        example: `-- Step 1: Create Audit Table
CREATE TABLE employee_audit (
    emp_id INT,
    action VARCHAR(20),
    changed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Step 2: Create Trigger Function
CREATE OR REPLACE FUNCTION log_emp_changes()
RETURNS TRIGGER AS $$
BEGIN
    INSERT INTO employee_audit(emp_id, action) 
    VALUES (NEW.emp_id, TG_OP);
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Step 3: Bind Trigger to Table
CREATE TRIGGER emp_after_insert
AFTER INSERT ON employees
FOR EACH ROW
EXECUTE FUNCTION log_emp_changes();`
      },
      {
        id: "exception-handling",
        title: "Exception Handling",
        definition: "Exception Handling is the block of code inside a PL/SQL program dedicated to trapping and handling runtime errors to prevent the entire transaction or script from crashing abruptly.",
        syntax: `BEGIN
    -- Normal block statements
EXCEPTION
    WHEN exception_name THEN
        -- Error response statements
END;`,
        typesOrFeatures: [
          "Predefined Exceptions: Automatically raised by the database engine for common SQL errors (e.g., NO_DATA_FOUND, ZERO_DIVIDE, DUP_VAL_ON_INDEX, TOO_MANY_ROWS, CURSOR_ALREADY_OPEN).",
          "User-Defined Exceptions: Custom errors declared by the developer and explicitly raised using the RAISE statement."
        ],
        advantages: [
          "Graceful Degradation: Allows transactions to rollback partially or print friendly error reports.",
          "Protects DB Integrity: Ensures errors are logged or compensated rather than raising raw application crashes."
        ],
        example: `CREATE OR REPLACE FUNCTION divide_salaries(emp_a INT, emp_b INT)
RETURNS NUMERIC AS $$
DECLARE
    sal_a NUMERIC;
    sal_b NUMERIC;
BEGIN
    SELECT salary INTO STRICT sal_a FROM employees WHERE emp_id = emp_a;
    SELECT salary INTO STRICT sal_b FROM employees WHERE emp_id = emp_b;
    
    RETURN sal_a / sal_b;
EXCEPTION
    WHEN ZERO_DIVIDE THEN
        RAISE WARNING 'Attempted to divide by zero! Returning 0.';
        RETURN 0;
    WHEN NO_DATA_FOUND THEN
        RAISE EXCEPTION 'One or both employee IDs do not exist!';
END;
$$ LANGUAGE plpgsql;`
      }
    ]
  },
  {
    id: "unit-3-transactions",
    title: "Unit 3: Transaction Management",
    description: "Transaction concepts, ACID properties, transaction state machine, serializability (conflict and view), and schedules.",
    topics: [
      {
        id: "acid-properties",
        title: "ACID Properties",
        definition: "ACID is an acronym representing the four critical properties of database transactions designed to guarantee data reliability, stability, and correctness even in the face of crashes or concurrent sessions.",
        typesOrFeatures: [
          "Atomicity ('All or Nothing'): Entire transaction completes successfully, or all its database updates are completely rolled back (undone). Managed by Recovery System.",
          "Consistency ('Preserving Rules'): Transforms the database from one valid state to another, obeying all integrity constraints (primary keys, foreign keys, etc.). Managed by Application/DBMS Schema.",
          "Isolation ('Independence'): Concurrent transactions run independently and must not witness each other's partial or uncommitted states. Managed by Concurrency Control.",
          "Durability ('Permanence'): Once a transaction commits, its changes are written to non-volatile storage and survive any subsequent system crashes. Managed by Log-Based Recovery."
        ],
        advantages: [
          "Ensures accurate bank-ledger operations (like bank transfers).",
          "Provides complete safety against operating system failures or server power cuts.",
          "Enables multiple applications to read and write to the same database simultaneously."
        ],
        example: `-- Classic Bank Transfer representing ACID
BEGIN;
UPDATE accounts SET balance = balance - 500 WHERE account_id = 'Alice';
UPDATE accounts SET balance = balance + 500 WHERE account_id = 'Bob';
-- If second step fails, the whole transaction rolls back (Atomicity)
COMMIT;`
      },
      {
        id: "transaction-states",
        title: "Transaction State Machine",
        definition: "A transaction goes through a sequence of distinct states during its lifecycle, from its initiation to its final termination (Commit or Abort).",
        typesOrFeatures: [
          "Active: Initial state where the transaction executes its read/write actions.",
          "Partially Committed: After the final statement has run, but before writing logs to disk.",
          "Failed: After the system discovers that normal execution can no longer proceed (error or deadlock).",
          "Aborted: The transaction is rolled back and database is restored to pre-transaction state.",
          "Committed: After successfully writing all log records to non-volatile storage. Changes are permanent!"
        ],
        advantages: [
          "Provides a systematic model for RDBMS design.",
          "Allows clean integration between transaction logs and recovery protocols."
        ],
        example: `
[Active] ---> [Partially Committed] ---> [Committed] (Success)
   |                 |
   v                 v
[Failed] --------> [Aborted] (Rollback/Cleanup)`
      },
      {
        id: "serializability",
        title: "Serializability & Schedules",
        definition: "Serializability is the standard criteria of correctness for concurrent transaction schedules. A concurrent schedule is serializable if its execution produces the exact same database state as a serial execution of those same transactions.",
        typesOrFeatures: [
          "Schedule: A sequence of operations (Read, Write, Commit, Abort) representing execution order of concurrent transactions.",
          "Conflict Serializability: A schedule is conflict serializable if it is conflict equivalent to some serial schedule. Checked via precedence graphs (cycles imply non-serializability).",
          "View Serializability: A more general model. Every conflict serializable schedule is view serializable, but some view serializable schedules (with blind writes) are not conflict serializable."
        ],
        advantages: [
          "Ensures consistency without forcing slow serial execution.",
          "Enables maximized hardware exploitation by running operations in parallel."
        ],
        example: `-- Precedence Graph Check:
-- If Transaction T1 writes data item X, and T2 subsequently reads X (W1(X) -> R2(X)),
-- there is a conflict. T1 must precede T2 in any serializable schedule (T1 -> T2).
-- If there is a cycle (T1 -> T2 -> T1), the schedule is NOT conflict serializable!`
      }
    ]
  },
  {
    id: "unit-4-concurrency",
    title: "Unit 4: Concurrency Control",
    description: "Locking protocols, Two-Phase Locking (2PL), Timestamp-based ordering, Thomas's Write Rule, and Validation protocols.",
    topics: [
      {
        id: "lock-based-protocols",
        title: "Lock-Based Protocols & Two-Phase Locking (2PL)",
        definition: "Locking protocols prevent concurrent access issues by requiring transactions to acquire a lock on a data item before reading or writing to it. Two-Phase Locking (2PL) is a specific lock-based protocol that guarantees conflict serializability.",
        syntax: `-- Locks can be requested explicitly in SQL:
SELECT * FROM accounts WHERE account_id = 'Alice' FOR UPDATE;`,
        typesOrFeatures: [
          "Shared Locks (S): Required for reading. Multiple transactions can hold shared locks on the same data item simultaneously.",
          "Exclusive Locks (X): Required for writing. Only one transaction can hold an exclusive lock on a data item; blocks all S and X requests.",
          "2PL Phase 1 (Growing Phase): A transaction can acquire locks but cannot release any.",
          "2PL Phase 2 (Shrinking Phase): A transaction can release locks but cannot acquire any new ones.",
          "Strict 2PL: Releases all exclusive locks only at the very end (Commit/Abort) to prevent cascading rollbacks."
        ],
        advantages: [
          "Guarantees conflict serializability.",
          "Strict 2PL prevents cascading rollbacks, ensuring highly recoverable schedules."
        ],
        example: `Transaction T1 (Growing Phase):
  Lock-S(A);  -- OK
  Lock-X(B);  -- OK
  -- (Point of maximum locks, Lock Point)
Transaction T1 (Shrinking Phase):
  Unlock(A);  -- Releases lock on A
  -- Cannot acquire any more locks now!
  Unlock(B);  -- Releases lock on B`
      },
      {
        id: "timestamp-based-ordering",
        title: "Timestamp-Based Ordering & Thomas's Write Rule",
        definition: "Timestamp-Based Ordering is a lock-free concurrency control protocol where the serialization order is determined entirely by assigning unique monotonic timestamps TS(T) to transactions when they initiate.",
        typesOrFeatures: [
          "Read Timestamp W-TS(Q): The largest timestamp of any transaction that successfully wrote to item Q.",
          "Read Timestamp R-TS(Q): The largest timestamp of any transaction that successfully read item Q.",
          "Timestamp ordering rules: If a transaction tries to read/write in a way that violates chronological order (e.g. reading a value that should have already been overwritten), it is rolled back and restarted.",
          "Thomas's Write Rule: An optimization of the timestamp protocol for blind writes. If a write operation is obsolete (older than the latest write timestamp W-TS(Q)), instead of rolling back, the write is simply ignored, and the transaction continues!"
        ],
        advantages: [
          "Deadlock-free: Since transactions never wait for locks, deadlocks are mathematically impossible.",
          "Highly suitable for read-heavy systems where conflicts are rare."
        ],
        example: `-- Thomas's Write Rule in Action:
-- Transaction T1 (TS = 10), Transaction T2 (TS = 20)
-- T2 writes to Q -> W-TS(Q) becomes 20.
-- T1 then tries to write to Q.
-- Normally: Rollback T1 because TS(T1) < W-TS(Q).
-- Thomas's Write Rule: Since Q is already written by T2 (younger), T1's write is obsolete.
-- Simply IGNORE T1's write and continue! No rollback needed!`
      }
    ]
  },
  {
    id: "unit-5-recovery",
    title: "Unit 5: Crash Recovery & Deadlocks",
    description: "Failure classifications, log-based recovery (deferred & immediate modification), checkpoints, restart recovery, and deadlock handling.",
    topics: [
      {
        id: "failure-classification",
        title: "Failure Classification & Storage Types",
        definition: "To design robust recovery systems, RDBMS classifies failures based on their scope and impact, dividing storage into volatile, non-volatile, and stable categories.",
        typesOrFeatures: [
          "Transaction Failure: Logical errors (user aborts) or system errors (division by zero).",
          "System Crash: Volatile RAM contents are lost due to power or OS failure, but disk data survives.",
          "Disk Failure: Physical head crash or bad sectors destroying database files on stable disk storage.",
          "Volatile Storage: Fast, loses data on power loss (RAM, cache).",
          "Non-Volatile Storage: Survives power loss (SSD, HDD, Magnetic tapes).",
          "Stable Storage: Theoretical storage that survives all physical failures (replicated across remote servers)."
        ],
        advantages: [
          "Allows structured recovery modeling.",
          "Directs backup schedules (regular backups for disk recovery, logs for system crash recovery)."
        ],
        example: `Crash Scenario:
1. Power fails. RAM goes black.
2. Recovery system boots up, reads the transaction log stored on disk (Non-Volatile),
3. Replays committed transactions (REDO) and rolls back uncommitted ones (UNDO).`
      },
      {
        id: "log-based-recovery",
        title: "Log-Based Recovery & Checkpoints",
        definition: "A log is a sequential append-only file of transaction records on disk. Log-based recovery mechanisms use this file to restore the database to a consistent state after a system crash.",
        syntax: `-- Simple representation of a log record
<T1, Start>
<T1, A, 1000, 800> -- T1 modified A from 1000 to 800
<T1, Commit>`,
        typesOrFeatures: [
          "Deferred Database Modification: Database updates are deferred to disk until the transaction commits. Log files record updates, but actual database disk files remain unchanged. Only needs REDO.",
          "Immediate Database Modification: Updates are written to disk directly while active. Needs both UNDO (to reverse uncommitted failures) and REDO.",
          "Checkpoints: Periodically, the RDBMS flushes all log records and dirty buffers to disk. This limits the size of the log that needs to be scanned during restart recovery."
        ],
        advantages: [
          "Guarantees atomicity and durability (ACID).",
          "Checkpoints vastly speed up database reboot times after crashes."
        ],
        example: `Recovery Process at Boot:
1. Locate the last checkpoint record in the log.
2. Maintain two lists: UNDO-list (uncommitted at crash) and REDO-list (committed).
3. Scan forward: Add transactions with <Ti, Commit> to REDO, others to UNDO.
4. Scan backward: Perform UNDO for all transactions in the UNDO list.
5. Scan forward: Perform REDO for all transactions in the REDO list.`
      },
      {
        id: "deadlocks",
        title: "Deadlock Handling (Prevention, Detection, Recovery)",
        definition: "A Deadlock is a system state where two or more concurrent transactions are stuck in a circular wait, each holding a lock that the other requires to proceed.",
        typesOrFeatures: [
          "Deadlock Prevention: Protocols to avoid deadlocks. Examples: Wait-Die scheme (older waits, younger dies) and Wound-Wait scheme (older wounds/kills younger, younger waits).",
          "Deadlock Detection: Periodically constructing a Wait-For Graph (nodes are transactions, directed edges indicate waiting for lock). Cycles indicate deadlocks.",
          "Deadlock Recovery: Selecting a victim transaction, rolling it back (partial or total) to release its locks, and restarting it."
        ],
        advantages: [
          "Ensures the system never hangs indefinitely.",
          "Balances fairness (preventing starvation) with system throughput."
        ],
        example: `Wait-Die Scheme:
- TS(T1) = 10, TS(T2) = 20 (T1 is older).
- T1 requests lock held by T2: T1 waits (T1 is older, so it waits).
- T2 requests lock held by T1: T2 dies (T2 is younger, so it dies/restarts).

Wound-Wait Scheme:
- T1 requests lock held by T2: T1 'wounds' T2, forcing T2 to abort/rollback.
- T2 requests lock held by T1: T2 waits.`
      }
    ]
  },
  {
    id: "unit-6-security",
    title: "Unit 6: Database Security & Access Control",
    description: "Database security vulnerabilities, roles, privileges, and relational SQL access control (GRANT, REVOKE).",
    topics: [
      {
        id: "db-security",
        title: "Database Security, Roles, and Privileges",
        definition: "Database security involves protecting data against unauthorized disclosure, accidental modification, or malicious destruction. This is managed via Discretionary Access Control (DAC) using database roles and privileges.",
        syntax: `CREATE ROLE role_name;
GRANT privilege_type ON object_name TO role_name;
REVOKE privilege_type ON object_name FROM role_name;`,
        typesOrFeatures: [
          "Roles: Named collections of privileges that can be granted to users or other roles. Greatly simplifies user administration.",
          "System Privileges: Rights to perform administrative tasks (e.g. CREATE TABLE, ALTER DATABASE).",
          "Object Privileges: Rights to execute operations on specific database schemas or objects (e.g., SELECT, INSERT, UPDATE, DELETE on a table).",
          "WITH GRANT OPTION: Allows a user or role to pass their granted privilege to other database users."
        ],
        advantages: [
          "Granular Security: Can restrict access down to specific columns or views.",
          "Role-Based Access Control (RBAC): Reduces error rates by grouping permissions according to job description.",
          "Easy auditing and tracking of data access patterns."
        ],
        example: `-- Create roles for university system
CREATE ROLE student;
CREATE ROLE registrar;

-- Grant permissions
GRANT SELECT ON courses TO student;
GRANT SELECT, INSERT, UPDATE, DELETE ON students, enrollments TO registrar;

-- Assign roles to database accounts
GRANT registrar TO alice;
GRANT student TO bob;

-- Revoking a privilege
REVOKE UPDATE ON students FROM registrar;`
      }
    ]
  }
];

export const EXPECTED_QUESTIONS: ExpectedQuestion[] = [
  {
    id: "eq-1",
    question: "Explain the ACID properties of database transactions in detail.",
    marks: 5,
    hint: "Think: All (Atomicity), Correct (Consistency), Independent (Isolation), Durable (Durability). For 5 marks, draw the states or list the component responsible for each.",
    idealOutline: {
      definition: "ACID is a set of properties that ensure database transactions are processed reliably. They guarantee relational sanity and database integrity during concurrent operations and crash events.",
      typesOrFeatures: [
        "Atomicity: 'All-or-nothing' execution. Handled by Transaction/Recovery manager.",
        "Consistency: Preserving correctness rules. Database must move from one valid state to another.",
        "Isolation: Concurrent sessions cannot read dirty data or uncommitted changes. Handled by Concurrency Control.",
        "Durability: Committed data is safe permanently on non-volatile disks. Handled by Log-based recovery."
      ],
      advantages: [
        "Prevents financial discrepancies.",
        "Ensures consistency during network drops.",
        "Enables stable multi-user environments."
      ],
      example: `BEGIN;
UPDATE accounts SET balance = balance - 100 WHERE id = 1;
UPDATE accounts SET balance = balance + 100 WHERE id = 2;
COMMIT;`
    }
  },
  {
    id: "eq-2",
    question: "What is a Trigger? Explain BEFORE, AFTER, and Row-level triggers with correct syntax and a realistic database audit example.",
    marks: 5,
    hint: "Use AFTER INSERT ROW level triggers to audit updates. Don't forget the return NEW; statement and the trigger function separation in PostgreSQL.",
    idealOutline: {
      definition: "A Trigger is a procedural program automatically invoked by the database engine in response to an DML operation (INSERT, UPDATE, DELETE) on a specified table.",
      syntax: `CREATE TRIGGER trigger_name
AFTER INSERT ON table_name
FOR EACH ROW EXECUTE FUNCTION audit_func();`,
      typesOrFeatures: [
        "BEFORE Triggers: Run before writing data. Used to validate or sanitize inputs.",
        "AFTER Triggers: Run after the data is successfully written. Used to log/audit activity.",
        "Row-level (FOR EACH ROW): Fires once per row affected.",
        "Statement-level: Fires once per query statement."
      ],
      advantages: [
        "Enforces bulletproof security auditing.",
        "Automates background calculations.",
        "Synchronizes backup stores."
      ],
      example: `CREATE OR REPLACE FUNCTION audit_log() RETURNS TRIGGER AS $$
BEGIN
    INSERT INTO audit_table(changed_by, log_time) VALUES (current_user, now());
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;`
    }
  },
  {
    id: "eq-3",
    question: "Explain Normalization. Differentiate between 1NF, 2NF, and 3NF with simple examples of functional dependencies.",
    marks: 5,
    hint: "Outline: 1NF = Atomic values. 2NF = 1NF + No partial dependencies. 3NF = 2NF + No transitive dependencies.",
    idealOutline: {
      definition: "Normalization is a systematic database structure design strategy that reduces redundancy and data anomalies by organizing relations into specific normal forms based on functional dependencies.",
      typesOrFeatures: [
        "1NF: Atomic values only, no repeating groups or multivalued arrays.",
        "2NF: No partial dependency (non-key columns must depend on the WHOLE primary key).",
        "3NF: No transitive dependency (non-key columns must not depend on other non-key columns)."
      ],
      advantages: [
        "Eliminates Insertion, Update, and Deletion anomalies.",
        "Minimizes structural database overhead.",
        "Prevents inconsistent duplicate states."
      ],
      example: `-- 3NF Corrected Split:
-- Table 1: Employee(EmpID, Name, DeptID)
-- Table 2: Department(DeptID, DeptName, Manager)`
    }
  },
  {
    id: "eq-4",
    question: "Explain %TYPE and %ROWTYPE with their advantages and write a PL/pgSQL block using them.",
    marks: 5,
    hint: "Binds types dynamically to table column or row structures. Write a clean DO block representing a SELECT INTO query.",
    idealOutline: {
      definition: "%TYPE and %ROWTYPE are procedural SQL attribute declarations used to copy the data type of database columns or tables dynamically into local block variables.",
      syntax: `DECLARE
    v_name employees.emp_name%TYPE;
    v_row employees%ROWTYPE;`,
      typesOrFeatures: [
        "%TYPE: Copies a single column's data type.",
        "%ROWTYPE: Copies an entire table's structure into a record variable.",
        "RECORD: Highly dynamic, structures itself on-the-fly."
      ],
      advantages: [
        "Maintains application compatibility when column types change.",
        "Reduces manual data-type matching overhead.",
        "Ensures high security and type safety."
      ],
      example: `DECLARE
    v_emp_id employees.emp_id%TYPE := 101;
    v_emp_record employees%ROWTYPE;
BEGIN
    SELECT * INTO v_emp_record FROM employees WHERE emp_id = v_emp_id;
    RAISE NOTICE 'Fetched Employee: %', v_emp_record.emp_name;
END;`
    }
  },
  {
    id: "eq-5",
    question: "Explain Two-Phase Locking (2PL) Protocol. How does it ensure Conflict Serializability?",
    marks: 5,
    hint: "Explain growing phase (acquires locks, releases none) and shrinking phase (releases locks, acquires none). Mention the Lock Point.",
    idealOutline: {
      definition: "The Two-Phase Locking (2PL) protocol is a concurrency control mechanism that guarantees conflict serializability by dividing a transaction's locking operations into two distinct sequential phases.",
      typesOrFeatures: [
        "Growing Phase: Transaction can acquire locks but cannot release any.",
        "Shrinking Phase: Transaction can release locks but cannot acquire any new ones.",
        "Lock Point: The exact moment when a transaction has acquired its final lock.",
        "Strict 2PL: All exclusive locks are held until commit/abort to avoid cascading rollback."
      ],
      advantages: [
        "Guarantees conflict serializable schedules.",
        "Prevents race conditions, dirty reads, and lost updates."
      ],
      example: `T1: Lock-X(A) -> Lock-S(B) -> [Lock Point] -> Unlock(A) -> Unlock(B)`
    }
  }
];

export const VIVA_QUESTIONS: VivaQuestion[] = [
  {
    id: "viva-1",
    question: "What is the difference between a Procedure and a Function in PL/SQL?",
    answer: "A Function must always return a value (or tabular set) using the RETURN statement, and can be called directly inside an SQL SELECT or WHERE query. A Procedure cannot be called inside a standard query; it executes a set of sequential steps, uses parameters (IN, OUT, INOUT) to pass multiple values back, and is executed using the CALL statement.",
    proTip: "Point out that Functions must have NO side-effects if used in SQL, and Procedures are preferred for complex database transaction sequences."
  },
  {
    id: "viva-2",
    question: "What are the special implicit cursor attributes and what do they represent?",
    answer: "In PostgreSQL, implicit cursor attributes can be checked after running an DML statement. For instance, FOUND returns true if at least one row was affected. In standard PL/SQL, attributes include: %FOUND (true if a row was fetched), %NOTFOUND (true if no rows left), %ROWCOUNT (returns number of fetched/affected rows), and %ISOPEN (always false for implicit cursors since they close automatically).",
    proTip: "Tell the examiner that you use 'EXIT WHEN NOT FOUND' or 'IF FOUND THEN' in your code, which shows practical familiarity."
  },
  {
    id: "viva-3",
    question: "What is a blind write and how is it handled by Thomas's Write Rule?",
    answer: "A blind write is a write operation on a data item without reading its value first. Thomas's Write Rule optimizes the timestamp ordering protocol for blind writes: if a transaction attempts to write to a data item that has already been overwritten by a younger transaction, instead of rolling back (which is the standard rule), the database simply ignores the write because it is obsolete, allowing execution to proceed smoothly.",
    proTip: "Mention that this is a rare protocol that improves performance by avoiding unnecessary transaction rollbacks."
  },
  {
    id: "viva-4",
    question: "What is a cascading rollback, and how do we prevent it?",
    answer: "A cascading rollback occurs when the failure of one transaction forces several other concurrent transactions to roll back because they read uncommitted data written by the failed transaction. We prevent this by enforcing Cascadeless Schedules, which are achieved by using Strict Two-Phase Locking (Strict 2PL), where exclusive locks are held until transaction commit.",
    proTip: "Summarize: 'Never let a transaction read uncommitted data from another transaction, and cascading rollback is physically impossible.'"
  },
  {
    id: "viva-5",
    question: "Explain the difference between GRANT and REVOKE commands.",
    answer: "GRANT is a DCL command used to provide specific database security permissions (such as SELECT, INSERT, UPDATE, or administrative roles) to a user or role. REVOKE is the inverse DCL command, used to withdraw those previously granted security privileges.",
    proTip: "Explain that REVOKE can propagate if the user granted permissions WITH GRANT OPTION, which demonstrates depth."
  }
];

export const MEMORY_TRICKS: MemoryTrick[] = [
  {
    id: "acid",
    concept: "ACID (Transaction Properties)",
    acronym: "ACID",
    expansion: "Atomicity, Consistency, Isolation, Durability",
    explanation: "Atomicity = All or nothing. Consistency = Integrity preserved. Isolation = Run in bubble. Durability = Survived crash.",
    color: "from-blue-600 to-indigo-600"
  },
  {
    id: "curt",
    concept: "SQL Command Categories",
    acronym: "DDL vs DML vs DCL vs TCL",
    expansion: "Definition, Manipulation, Control, Transaction Control",
    explanation: "DDL (Create, Alter, Drop). DML (Select, Insert, Update, Delete). DCL (Grant, Revoke). TCL (Commit, Rollback, Savepoint).",
    color: "from-emerald-600 to-teal-600"
  },
  {
    id: "crud",
    concept: "Basic DB Operations",
    acronym: "CRUD",
    expansion: "Create, Read, Update, Delete",
    explanation: "The four basic operations of persistent storage. Map to SQL: INSERT, SELECT, UPDATE, DELETE.",
    color: "from-amber-600 to-orange-600"
  },
  {
    id: "two-pl",
    concept: "2PL Phases",
    acronym: "G-S (Growing & Shrinking)",
    expansion: "Growing Phase (Acquire), Shrinking Phase (Release)",
    explanation: "Once you release a lock, you cannot acquire any more locks! Growing = only lock. Shrinking = only unlock.",
    color: "from-rose-600 to-pink-600"
  },
  {
    id: "wait-for",
    concept: "Deadlock schemes",
    acronym: "W-D vs W-W",
    expansion: "Wait-Die (Older waits, Younger dies) vs Wound-Wait (Older wounds, Younger waits)",
    explanation: "Wait-Die is non-preemptive (older waits). Wound-Wait is preemptive (older preempts/wounds younger).",
    color: "from-purple-600 to-violet-600"
  }
];
