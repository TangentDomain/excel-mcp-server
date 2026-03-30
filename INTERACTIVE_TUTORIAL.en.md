# 🎮 ExcelMCP Interactive Tutorial

[English](INTERACTIVE_TUTORIAL.en.md) | [简体中文](INTERACTIVE_TUTORIAL.md)

> **Game Development Excel Configuration Management - Interactive Learning Guide**  
> From zero to hero, learn by doing, master ExcelMCP core features quickly

---

## 🚀 Quick Start

### Step 1: Installation & Setup
```bash
# One-click installation
pip install excel-mcp-server-fastmcp

# Verify installation
excel-mcp-server-fastmcp --version
```

**✅ Completion Check**: You should see version output like `v1.6.49`

### Step 2: First Launch
```bash
# Start MCP server
excel-mcp-server-fastmcp

# Expected output:
# 🎮 ExcelMCP Game Development Excel Manager v1.6.49
# 📊 53 tools loaded and ready
# 💡 Type "help" to see available tools
```

**✅ Completion Check**: You should see the startup success message

---

## 🎯 Basic Skills Training

### Skill 1: Data Query (5 minutes)

#### Objective
Learn to query Excel data using natural language

#### Exercise 1: Simple Query
**Task**: Query all data from the skills table
**Input Command**:
```text
Query all content from the skills table
```

**Expected Result**:
- Shows all columns from skills table
- Includes skill ID, name, type, cost, etc.

#### Exercise 2: Conditional Query
**Task**: Find all mage skills
**Input Command**:
```text
Find all skills with type 'mage' in the skills table
```

**Expected Result**:
- Shows only mage-type skills
- Includes complete skill information

#### ✅ Verification Checklist
- [ ] Can execute simple queries successfully
- [ ] Can execute conditional queries successfully
- [ ] Understand the query result format

---

### Skill 2: Data Update (5 minutes)

#### Objective
Learn to batch update Excel data

#### Exercise 1: Value Adjustment
**Task**: Increase damage of all attack skills by 10%
**Input Command**:
```text
Increase the damage of all attack skills in the skills table by 10%
```

**Expected Result**:
- Attack skills' damage increased by 10%
- Other skills data unchanged

#### Exercise 2: Attribute Modification
**Task**: Reduce MP cost of all mage skills by 5%
**Input Command**:
```text
Reduce the MP cost of all mage skills in the skills table by 5%
```

**Expected Result**:
- Mage skills' MP cost reduced by 5%
- Other skill attributes unchanged

#### ✅ Verification Checklist
- [ ] Can execute numerical updates successfully
- [ ] Can execute conditional updates successfully
- [ ] Understand batch update safety

---

### Skill 3: Cross-table Operations (8 minutes)

#### Objective
Learn to use JOIN queries to relate multiple tables

#### Scenario Setup
Assume you have the following tables:
- **Skills Table**: Skill ID, Skill Name, Skill Type, MP Cost, Damage Value
- **Classes Table**: Class ID, Class Name, Main Skill Type, Difficulty Level

#### Exercise 1: JOIN Query
**Task**: Query each class and its main skill
**Input Command**:
```text
Query the classes table and skills table to find each class's main skill, sorted by class difficulty
```

**Expected Result**:
- Shows class name + corresponding main skill
- Sorted by difficulty level

#### Exercise 2: Complex Analysis
**Task**: Analyze skill costs by class
**Input Command**:
```text
Analyze the classes table and skills table to calculate average skill cost per class, find the class with highest cost
```

**Expected Result**:
- Shows average cost per class
- Highlights the class with highest cost

#### ✅ Verification Checklist
- [ ] Can execute JOIN queries successfully
- [ ] Can understand related data meaning
- [ ] Can perform cross-table data analysis

---

## 🎮 Game Scenario实战

### Scenario 1: Skill System Optimization (10 minutes)

#### Background
You're developing an RPG game and need to balance the skill system:

- Skills table includes: Attack skills, Defense skills, Healing skills
- Current issue: Attack skills are too strong, need adjustment

#### Task List
1. **Current Analysis**
   ```text
   Analyze the skills table, count skills by type and calculate average damage
   ```

2. **Balance Adjustment**
   ```text
   Reduce all attack skills' damage by 15%, increase all defense skills' duration by 20%
   ```

3. **Effect Verification**
   ```text
   Re-analyze the skills table to verify balance after adjustments
   ```

#### ✅ Completion Standard
- [ ] Can analyze skill distribution
- [ ] Can successfully adjust skill balance
- [ ] Can verify adjustment effects

---

### Scenario 2: Equipment Configuration Management (10 minutes)

#### Background
Game has大量 equipment that needs configuration management:

- Equipment table includes: Equipment Name, Rarity, Attack Power, Defense Power, Price
- Need to manage by rarity category

#### Task List
1. **Equipment Classification**
   ```text
   Classify equipment by rarity, count equipment and calculate average price per rarity
   ```

2. **Price Adjustment**
   ```text
   Increase price of all rare equipment by 20%, keep common equipment prices unchanged
   ```

3. **Set Detection**
   ```text
   Detect set equipment (equipment with "set" in name)
   ```

#### ✅ Completion Standard
- [ ] Can classify and count by rarity
- [ ] Can batch adjust equipment prices
- [ ] Can detect set equipment

---

## 🚀 Advanced Skill Challenges

### Challenge 1: Complex Query Optimization

#### Task
Query skill effects table to find:
1. Attack skills with damage > 100
2. Healing skills with MP cost < 50
3. Sort by damage/cost ratio to find highest value skills

#### Command Hint
```text
Query skill effects table, find attack skills with damage>100 and healing skills with cost<50, sort by damage cost ratio descending
```

#### ✅ Challenge Completion
- [ ] Can query multiple conditions simultaneously
- [ ] Can perform complex sorting
- [ ] Can calculate derived metrics

---

### Challenge 2: Data Consistency Check

#### Task
Check consistency between equipment table and inventory table:
1. Find equipment in equipment table but not in inventory
2. Find items in inventory but not in equipment table
3. Generate difference report

#### Command Hint
```text
Compare equipment ID between equipment table and inventory table, find inconsistent items, generate difference report
```

#### ✅ Challenge Completion
- [ ] Can check data consistency
- [ ] Can identify data differences
- [ ] Can generate analysis reports

---

## 💡 Learning Tips

### 🎯 Effective Practice Methods
1. **Start Simple**: Master basic queries first, then try complex operations
2. **Real Scenarios**: Practice with real game data for better results
3. **Progressive Challenges**: Complete basic exercises before attempting advanced challenges
4. **Regular Review**: Review and reinforce skills regularly

### 🔍 Common Problem Solutions
- **No Query Results**: Check table and field names are correct
- **Update Failed**: Confirm data types and formats are correct
- **Performance Issues**: Use pagination for large files, avoid full loading

### 📞 Getting Help
- **View Help**: Type "help" to see all available tools
- **Check Logs**: See detailed logs in `.excel_mcp_logs/` directory
- **Report Issues**: [GitHub Issues](https://github.com/TangentDomain/excel-mcp-server/issues/new)

---

## 🎉 Congratulations Complete!

Through the above exercises, you've mastered ExcelMCP core skills:

### ✅ Skills Mastered
- ✅ Basic Data Queries
- ✅ Batch Data Updates
- ✅ Cross-table JOIN Operations
- ✅ Game Scene Applications
- ✅ Complex Query Optimization

### 🚀 Next Steps
1. **Real Project Application**: Use ExcelMCP in your own game projects
2. **Advanced Features**: Try CTE queries, subqueries, and other advanced features
3. **Community Participation**: Share your experience and help other developers
4. **Feature Feedback**: Provide usage feedback to help improve the product

**Keep going, let Excel configuration management become your development accelerator!** 🚀

---
*Tutorial Version: v1.6.50*  
*Updated: 2026-03-29*  
*Suggested Learning Time: 30-45 minutes*