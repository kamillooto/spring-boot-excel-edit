package com.kamillooto.springbootexceledit.springbootexceledit;

import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.List;

@Configuration
public class ProductConfig {
    @Bean
    CommandLineRunner commandLineRunner(ProductRepository productRepository) {
        return args -> {
            Product product1 = new Product(
                    1,
                    "product1",
                    1000
            );
            Product product2 = new Product(
                    2,
                    "product2",
                    2000
            );
            Product product3 = new Product(
                    3,
                    "product3",
                    3000
            );
            productRepository.saveAll(
                    List.of(product1, product2, product3)
            );
        };
    }
}
