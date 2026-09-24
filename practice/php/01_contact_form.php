<?php
// Exercise 1: Contact form with server side validation
//
// Run with:  php -S localhost:8000   then open http://localhost:8000/01_contact_form.php
// Topics: $_POST, validation, htmlspecialchars (XSS protection), sticky form values

function clean($value)
{
    return htmlspecialchars(trim($value), ENT_QUOTES, 'UTF-8');
}

function validate($data)
{
    $errors = [];

    if (strlen(trim($data['name'] ?? '')) < 2) {
        $errors['name'] = 'Name must have at least 2 characters.';
    }
    if (!filter_var($data['email'] ?? '', FILTER_VALIDATE_EMAIL)) {
        $errors['email'] = 'Please enter a valid email address.';
    }
    $allowed = ['general', 'support', 'other'];
    if (!in_array($data['topic'] ?? '', $allowed, true)) {
        $errors['topic'] = 'Please choose a topic.';
    }
    if (strlen(trim($data['message'] ?? '')) < 10) {
        $errors['message'] = 'Message must have at least 10 characters.';
    }

    return $errors;
}

$errors = [];
$sent = false;
$values = ['name' => '', 'email' => '', 'topic' => '', 'message' => ''];

if (($_SERVER['REQUEST_METHOD'] ?? 'GET') === 'POST') {
    foreach ($values as $key => $_) {
        $values[$key] = $_POST[$key] ?? '';
    }
    $errors = validate($values);

    if (empty($errors)) {
        // Save the message to a text file instead of sending a real email
        $line = date('Y-m-d H:i') . ' | ' . clean($values['name']) . ' | ' . clean($values['email'])
            . ' | ' . $values['topic'] . ' | ' . str_replace("\n", ' ', clean($values['message'])) . PHP_EOL;
        file_put_contents(__DIR__ . '/messages.txt', $line, FILE_APPEND);
        $sent = true;
        $values = ['name' => '', 'email' => '', 'topic' => '', 'message' => ''];
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Contact form</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 500px; margin: 40px auto; }
        label { display: block; margin-top: 12px; }
        input, select, textarea { width: 100%; padding: 6px; box-sizing: border-box; }
        .error { color: #b00020; font-size: 0.9em; }
        .success { background: #e6f4ea; padding: 10px; border-radius: 4px; }
        button { margin-top: 16px; padding: 8px 16px; }
    </style>
</head>
<body>
<h1>Contact us</h1>

<?php if ($sent): ?>
    <p class="success">Thank you, your message was saved.</p>
<?php endif; ?>

<form method="post" novalidate>
    <label>Name
        <input type="text" name="name" value="<?= clean($values['name']) ?>">
    </label>
    <?php if (isset($errors['name'])): ?><div class="error"><?= $errors['name'] ?></div><?php endif; ?>

    <label>Email
        <input type="email" name="email" value="<?= clean($values['email']) ?>">
    </label>
    <?php if (isset($errors['email'])): ?><div class="error"><?= $errors['email'] ?></div><?php endif; ?>

    <label>Topic
        <select name="topic">
            <option value="">-- choose --</option>
            <?php foreach (['general' => 'General question', 'support' => 'Support', 'other' => 'Other'] as $key => $label): ?>
                <option value="<?= $key ?>" <?= $values['topic'] === $key ? 'selected' : '' ?>><?= $label ?></option>
            <?php endforeach; ?>
        </select>
    </label>
    <?php if (isset($errors['topic'])): ?><div class="error"><?= $errors['topic'] ?></div><?php endif; ?>

    <label>Message
        <textarea name="message" rows="5"><?= clean($values['message']) ?></textarea>
    </label>
    <?php if (isset($errors['message'])): ?><div class="error"><?= $errors['message'] ?></div><?php endif; ?>

    <button type="submit">Send</button>
</form>
</body>
</html>
