package com.generate.xlsx.spring.boot.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.generate.xlsx.spring.boot.model.Role;

@Repository
public interface RoleRepository extends JpaRepository<Role, Integer> {

}