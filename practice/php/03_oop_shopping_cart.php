<?php
// Exercise 3: Shopping cart with classes and interfaces
//
// Run with:  php 03_oop_shopping_cart.php
// Topics: classes, interfaces, constructor promotion, exceptions, array functions

interface Discount
{
    public function apply(float $total): float;
    public function describe(): string;
}

class PercentDiscount implements Discount
{
    public function __construct(private float $percent)
    {
        if ($percent <= 0 || $percent >= 100) {
            throw new InvalidArgumentException('Percent must be between 0 and 100');
        }
    }

    public function apply(float $total): float
    {
        return $total * (1 - $this->percent / 100);
    }

    public function describe(): string
    {
        return "{$this->percent}% off";
    }
}

class FixedDiscount implements Discount
{
    public function __construct(private float $amount, private float $minimumTotal = 0)
    {
    }

    public function apply(float $total): float
    {
        return $total >= $this->minimumTotal ? max(0, $total - $this->amount) : $total;
    }

    public function describe(): string
    {
        return "{$this->amount} KM off orders over {$this->minimumTotal} KM";
    }
}

class Product
{
    public function __construct(public string $sku, public string $name, public float $price)
    {
    }
}

class Cart
{
    /** @var array<string, array{product: Product, qty: int}> */
    private array $items = [];
    private ?Discount $discount = null;

    public function add(Product $product, int $qty = 1): void
    {
        if ($qty < 1) {
            throw new InvalidArgumentException('Quantity must be at least 1');
        }
        if (isset($this->items[$product->sku])) {
            $this->items[$product->sku]['qty'] += $qty;
        } else {
            $this->items[$product->sku] = ['product' => $product, 'qty' => $qty];
        }
    }

    public function remove(string $sku): void
    {
        unset($this->items[$sku]);
    }

    public function setDiscount(Discount $discount): void
    {
        $this->discount = $discount;
    }

    public function subtotal(): float
    {
        return array_sum(array_map(fn($i) => $i['product']->price * $i['qty'], $this->items));
    }

    public function total(): float
    {
        $subtotal = $this->subtotal();
        return round($this->discount ? $this->discount->apply($subtotal) : $subtotal, 2);
    }

    public function printReceipt(): void
    {
        foreach ($this->items as $item) {
            printf("%-12s %3d x %7.2f = %8.2f\n",
                $item['product']->name, $item['qty'], $item['product']->price,
                $item['product']->price * $item['qty']);
        }
        echo str_repeat('-', 38) . "\n";
        printf("%-27s %10.2f\n", 'Subtotal', $this->subtotal());
        if ($this->discount) {
            printf("%-27s %10s\n", 'Discount', $this->discount->describe());
        }
        printf("%-27s %10.2f\n", 'TOTAL', $this->total());
    }
}

$cart = new Cart();
$cart->add(new Product('P1', 'Keyboard', 45.00));
$cart->add(new Product('P2', 'Mouse', 19.90), 2);
$cart->add(new Product('P3', 'USB cable', 5.50), 3);
$cart->add(new Product('P2', 'Mouse', 19.90));   // adds to the existing quantity
$cart->remove('P3');

$cart->setDiscount(new PercentDiscount(10));
$cart->printReceipt();

echo "\n";
$cart->setDiscount(new FixedDiscount(20, 100));
$cart->printReceipt();

try {
    $cart->add(new Product('P4', 'Monitor', 250), 0);
} catch (InvalidArgumentException $e) {
    echo "\nError: " . $e->getMessage() . "\n";
}
