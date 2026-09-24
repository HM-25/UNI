<?php
// Exercise 2: Students CRUD with PDO and SQLite
//
// Run with:  php -S localhost:8000   then open http://localhost:8000/02_students_crud.php
// Topics: PDO, prepared statements (SQL injection protection), CRUD, GET/POST

$db = new PDO('sqlite:' . __DIR__ . '/students.db');
$db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$db->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);

$db->exec('CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    first_name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    index_number TEXT NOT NULL UNIQUE,
    year INTEGER NOT NULL CHECK (year BETWEEN 1 AND 5)
)');

function e($value)
{
    return htmlspecialchars((string)$value, ENT_QUOTES, 'UTF-8');
}

$error = '';
$action = $_POST['action'] ?? '';

try {
    if ($action === 'create' || $action === 'update') {
        $data = [
            ':first' => trim($_POST['first_name'] ?? ''),
            ':last'  => trim($_POST['last_name'] ?? ''),
            ':index' => trim($_POST['index_number'] ?? ''),
            ':year'  => (int)($_POST['year'] ?? 0),
        ];
        if ($data[':first'] === '' || $data[':last'] === '' || $data[':index'] === '') {
            throw new InvalidArgumentException('All fields are required.');
        }

        if ($action === 'create') {
            $stmt = $db->prepare('INSERT INTO students (first_name, last_name, index_number, year)
                                  VALUES (:first, :last, :index, :year)');
        } else {
            $data[':id'] = (int)$_POST['id'];
            $stmt = $db->prepare('UPDATE students SET first_name = :first, last_name = :last,
                                  index_number = :index, year = :year WHERE id = :id');
        }
        $stmt->execute($data);
        header('Location: ' . strtok($_SERVER['REQUEST_URI'] ?? '', '?'));
        exit;
    }

    if ($action === 'delete') {
        $stmt = $db->prepare('DELETE FROM students WHERE id = ?');
        $stmt->execute([(int)$_POST['id']]);
    }
} catch (PDOException $ex) {
    $error = str_contains($ex->getMessage(), 'UNIQUE') ? 'Index number already exists.' : 'Database error.';
} catch (InvalidArgumentException $ex) {
    $error = $ex->getMessage();
}

// Edit mode: load one student into the form
$edit = null;
if (isset($_GET['edit'])) {
    $stmt = $db->prepare('SELECT * FROM students WHERE id = ?');
    $stmt->execute([(int)$_GET['edit']]);
    $edit = $stmt->fetch() ?: null;
}

// Search
$search = trim($_GET['q'] ?? '');
if ($search !== '') {
    $stmt = $db->prepare('SELECT * FROM students
                          WHERE first_name LIKE :q OR last_name LIKE :q OR index_number LIKE :q
                          ORDER BY last_name');
    $stmt->execute([':q' => "%$search%"]);
} else {
    $stmt = $db->query('SELECT * FROM students ORDER BY last_name');
}
$students = $stmt->fetchAll();
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Students</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 30px auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
        th { background: #f0f0f0; }
        form.inline { display: inline; }
        .error { color: #b00020; }
        fieldset { margin-top: 20px; }
    </style>
</head>
<body>
<h1>Students</h1>

<form method="get">
    <input type="text" name="q" placeholder="Search..." value="<?= e($search) ?>">
    <button>Search</button>
</form>

<?php if ($error): ?><p class="error"><?= e($error) ?></p><?php endif; ?>

<table>
    <tr><th>Index</th><th>First name</th><th>Last name</th><th>Year</th><th></th></tr>
    <?php foreach ($students as $s): ?>
        <tr>
            <td><?= e($s['index_number']) ?></td>
            <td><?= e($s['first_name']) ?></td>
            <td><?= e($s['last_name']) ?></td>
            <td><?= e($s['year']) ?></td>
            <td>
                <a href="?edit=<?= $s['id'] ?>">Edit</a>
                <form method="post" class="inline" onsubmit="return confirm('Delete this student?')">
                    <input type="hidden" name="action" value="delete">
                    <input type="hidden" name="id" value="<?= $s['id'] ?>">
                    <button>Delete</button>
                </form>
            </td>
        </tr>
    <?php endforeach; ?>
    <?php if (!$students): ?><tr><td colspan="5">No students found.</td></tr><?php endif; ?>
</table>

<fieldset>
    <legend><?= $edit ? 'Edit student' : 'Add student' ?></legend>
    <form method="post">
        <input type="hidden" name="action" value="<?= $edit ? 'update' : 'create' ?>">
        <?php if ($edit): ?><input type="hidden" name="id" value="<?= $edit['id'] ?>"><?php endif; ?>
        <input name="first_name" placeholder="First name" value="<?= e($edit['first_name'] ?? '') ?>">
        <input name="last_name" placeholder="Last name" value="<?= e($edit['last_name'] ?? '') ?>">
        <input name="index_number" placeholder="Index number" value="<?= e($edit['index_number'] ?? '') ?>">
        <select name="year">
            <?php for ($y = 1; $y <= 5; $y++): ?>
                <option <?= (int)($edit['year'] ?? 1) === $y ? 'selected' : '' ?>><?= $y ?></option>
            <?php endfor; ?>
        </select>
        <button><?= $edit ? 'Save' : 'Add' ?></button>
        <?php if ($edit): ?><a href="?">Cancel</a><?php endif; ?>
    </form>
</fieldset>
</body>
</html>
