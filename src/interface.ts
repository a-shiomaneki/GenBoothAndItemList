export interface Factory {
    create(...args: any[]): Product;
    createProduct(...args: any[]): Product;
    registerProduct(product: Product): void;
}

export interface Product {
    filename: string;
}

